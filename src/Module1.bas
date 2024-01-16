Attribute VB_Name = "Module1"
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
    
    For Each Sht In ActiveWorkbook.Worksheets
        If Sht.Name <> shtActive.Name Then
            Sht.Select
            Set rngData = Sht.Range("A1").CurrentRegion
            With rngData
                Set rngLastRowCells = .Rows(.Rows.count)
                .HorizontalAlignment = xlCenter
                .Font.Color = vbWhite
'                varArrColumnTitles = Array("confirmed_fraud_sessions", "total_sessions", "confirmed_fraud_puids", "total_puids")
'                For i = 0 To UBound(varArrColumnTitles)
'                    Cells(rngLastRowCells.Row, GetSheetColumnIndexByTitle(CStr(varArrColumnTitles(i)), Sht)).FormulaR1C1 = "=SUM(R2C:R[-1]C)"
'                Next i
            
'                varArrColumnTitles = Array("session_precision_prc", "puid_precision_prc")
'                For i = 0 To UBound(varArrColumnTitles)
'                    Cells(rngLastRowCells.Row, GetSheetColumnIndexByTitle(CStr(varArrColumnTitles(i)), Sht)).FormulaR1C1 = "=RC[-2]/RC[-1]"
'                    Sht.Columns(GetSheetColumnIndexByTitle(CStr(varArrColumnTitles(i)), Sht)).Style = "Percent"
'                Next i
            
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



