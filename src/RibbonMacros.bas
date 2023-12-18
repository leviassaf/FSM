Attribute VB_Name = "RibbonMacros"
Option Explicit

Public myRibbon As IRibbonUI

Sub Onload(ribbon As IRibbonUI)
    'Create a ribbon instance for use in this project
    Set myRibbon = ribbon
End Sub

'DropDown onAction
Sub runReport(ByVal control As IRibbonControl, selectedID As String, selectedIndex As Integer)
    On Error Resume Next
    Dim strList As String
    Dim strMacro As String
    Dim lastRow As Long
    Dim dtStart As Date
    Dim dtEnd As Date
    
    If selectedIndex = 0 Then
        Exit Sub
    End If
    
    lastRow = Log.Cells(Log.Rows.count, "A").End(xlUp).Row + 1
    
    Select Case control.id
    Case "TrustboardReports"
        strList = "TrustboardReports"
    Case "AutobotReports"
        strList = "AutobotReports"
    Case "SisenseReports"
        strList = "SisenseReports"
    Case "DbReports"
Stop
        strList = "SQLQueriesNames"
    End Select
    
    strMacro = ThisWorkbook.Names(strList).RefersToRange.Rows(CLng(selectedIndex + 1)).Value
    
    With Log
        .Cells(lastRow, 1).Value = strMacro
        dtStart = Now
        .Cells(lastRow, 2).Value = dtStart
        .Cells(lastRow, 5).Value = Application.username
    End With
    
    Application.Run (strMacro)
    
    With Log
        dtEnd = Now
        .Cells(lastRow, 3).Value = dtEnd
        .Cells(lastRow, 4).Value = dtEnd - dtStart
    End With
    'Restore control to original state
    myRibbon.InvalidateControl control.id
End Sub

'Callback for DropDown getItemCount
Sub GetItemCount(ByVal control As IRibbonControl, ByRef count)
    Dim strList As String
    
    Select Case control.id
    Case "TrustboardReports"
        strList = "TrustboardReports"
    Case "AutobotReports"
        strList = "AutobotReports"
    Case "SisenseReports"
        strList = "SisenseReports"
    Case "DbReports"
'Stop
        strList = "SQLQueriesNames"
    End Select
    count = ThisWorkbook.Names(strList).RefersToRange.Rows.count
End Sub

'Callback for DropDown getItemLabel
Sub GetItemLabel(ByVal control As IRibbonControl, index As Integer, ByRef label)
    Dim rngML As Range
    Dim strList As String
    
    Select Case control.id
    Case "TrustboardReports"
        strList = "TrustboardReports"
    Case "AutobotReports"
        strList = "AutobotReports"
    Case "SisenseReports"
        strList = "SisenseReports"
    Case "DbReports"
'Stop
        strList = "SQLQueriesNames"
    End Select
    
    Set rngML = ThisWorkbook.Names(strList).RefersToRange
    label = rngML.Cells(index + 1)
End Sub

'Callback for DropDown getSelectedItemIndex
Sub GetSelItemIndex(ByVal control As IRibbonControl, ByRef index)
    'Ensure first item in dropdown is displayed.
    Select Case control.id
        Case Is = "TrustboardReports"
        index = 0
    Case Is = "AutobotReports"
        index = 0
    Case Is = "SisenseReports"
        index = 0
    Case Is = "DbReports"
'Stop
        index = 0
    Case Else
    End Select
End Sub
