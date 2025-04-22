Attribute VB_Name = "modFormatPolicyPerformance"
Option Explicit

Sub FormatPolicyPerformance()
    Dim rngData As Range
    Dim shtActive As Worksheet
    Dim varArrColumnTitles As Variant
    Dim i As Integer
    Dim rngLastRowCells As Range
    Dim Sht As Worksheet
    
    Set shtActive = ActiveSheet
    
    Call SplitWorksheetsByColumnValues("calendar_week", ActiveSheet)
    
    ActiveWorkbook.ApplyTheme ( _
        "C:\Program Files\Microsoft Office\root\Document Themes 16\Office 2013 - 2022 Theme.thmx")

    For Each Sht In ActiveWorkbook.Worksheets
        If Sht.name <> shtActive.name Then
            Sht.Select
            Set rngData = Sht.Range("A1").CurrentRegion
            With rngData
                Set rngLastRowCells = .Rows(.Rows.count)
                .HorizontalAlignment = xlCenter
                .Font.Color = vbWhite
            
                varArrColumnTitles = Array("confirmed_fraud_sessions", "confirmed_fraud_puids")
                For i = 0 To UBound(varArrColumnTitles)
                    Intersect(rngData, Sht.Columns(GetSheetColumnIndexByTitle(CStr(varArrColumnTitles(i)), Sht, Sht.Range("A1")))).Font.Color = vbGreen
                Next i
                
                rngLastRowCells.Font.Bold = True
            
                Sht.Cells(rngLastRowCells.Row, 3).Value = "Grand Average/Total"
                Sht.Range(Cells(rngLastRowCells.Row, 3), Cells(rngLastRowCells.Row, 4)).Merge
            
                .Interior.ThemeColor = xlThemeColorLight1
            End With
        End If
    Next Sht
End Sub



