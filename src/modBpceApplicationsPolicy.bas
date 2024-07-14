Attribute VB_Name = "modBpceApplicationsPolicy"
Option Explicit

Public Const MAX_SHEET_NAME_LENGTH As Integer = 31

Sub BpceApplicationsPolicy()
    Dim Wbk As Workbook
    Dim shtRawData As Worksheet
    Dim strDetectionRateFolderPath As String
    Dim strBoxPath As String
    Dim arrColumnsWithExceptions As Variant
    Dim Pvt As PivotTable
    Const REPORT_NAME As String = "Pivot Table"
    Dim intNumberOfSourceFiles As Integer
    Dim strQueryString As String
    Dim arrVarValuesRiskReasons() As Variant
    Dim strFileName As String
    Dim shtPivot As Worksheet
    Dim shtApplicationPivot As Worksheet
    Dim strApplicationName As String
    
'    strDetectionRateFolderPath = GetDownloadsPath & Application.PathSeparator & "bpce temp"
    strDetectionRateFolderPath = getSelectedFolder(strBoxPath)
    If strDetectionRateFolderPath = "False" Then Exit Sub
    Application.ScreenUpdating = False
    intNumberOfSourceFiles = CountFilesInFolder(strDetectionRateFolderPath)
    strQueryString = "let" & vbNewLine & "    Source = Folder.Files(""" & strDetectionRateFolderPath & """)," & vbNewLine & _
        "    #""Filtered Hidden Files1"" = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true)," & vbNewLine & _
        "    #""Invoke Custom Function1"" = Table.AddColumn(#""Filtered Hidden Files1"", ""Transform File"", each #""Transform File""([Content]))," & vbNewLine & _
        "    #""Renamed Columns1"" = Table.RenameColumns(#""Invoke Custom Function1"", {""Name"", ""Source.Name""})," & vbNewLine & _
        "    #""Removed Other Columns1"" = Table.SelectColumns(#""Renamed Columns1"", {""Source.Name"", ""Transform File""})," & vbNewLine & _
        "    #""Expanded Table Column1"" = Table.ExpandTableColumn(#""Removed Other Columns1"", ""Transform File"", Table.ColumnNames(#""Transform File""(#""Sample File"")))," & vbNewLine & _
        "    #""Changed Type"" = Table.TransformColumnTypes(#""Expanded Table Column1"",{{""Source.Name"", type text}, {""policy_category"", type text}, {""activity"", type text}, {""rule_name"", type text}, {""score"", Int64.Type}, {""reason_id"", type text}, {""recommendation"", type text}, {""rule_id"", Int64.Type}})" & vbNewLine & _
        "in" & vbNewLine & _
        "    #""Changed Type"""

    Set Wbk = Workbooks.Add(xlWBATWorksheet)
    With Wbk
        With .Queries
            If intNumberOfSourceFiles > 1 Then 'if more than 1 source file was found
                .Add Name:="foo report name", _
                    Formula:=strQueryString
                .Add Name:="Sample File", Formula:= _
                    "let Source = Folder.Files(""" & strDetectionRateFolderPath & """), Navigation1 = Source{0}[Content] in Navigation1"
                .Add Name:="Parameter1", Formula:= _
                    "#""Sample File"" meta [IsParameterQuery=true, BinaryIdentifier=#""Sample File"", Type=""Binary"", IsParameterQueryRequired=true]"
                .Add Name:="Transform Sample File", Formula:= _
                    "let Source = Csv.Document(Parameter1,[Delimiter="","", Columns=7, QuoteStyle=QuoteStyle.None]), #""Promoted Headers"" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]) in #""Promoted Headers"""
                .Add Name:="Transform File", Formula:= _
                    "let Source = (Parameter1) => let Source = Csv.Document(Parameter1,[Delimiter="","", Columns=7, QuoteStyle=QuoteStyle.None]), #""Promoted Headers"" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]) in #""Promoted Headers"" in Source"
            Else
                MsgBox "Adjust VBA code to handle importing a single source file"
            End If
        End With
        Set shtRawData = .ActiveSheet
        With shtRawData
            With .ListObjects.Add(SourceType:=0, Source:= _
            "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""foo report name"";Extended Properties=""""" _
            , Destination:=Range("$A$1")).QueryTable
            .CommandType = xlCmdSql
            .CommandText = Array("SELECT * FROM [foo report name]")
            .Refresh BackgroundQuery:=False
            End With
            .Name = "Raw Data"
        End With
        Call ReplaceCsvExtensionWithNone
        
        Call ActiveReasonsPerActivityInternal
        
        Set shtPivot = ActiveSheet
        Set Pvt = shtPivot.PivotTables(1)
        With Pvt
            With .PivotFields("Source.Name")
                .Orientation = xlPageField
                .Position = 1
            End With
            
            .ShowPages PageField:="Source.Name"
            Application.DisplayAlerts = False
            shtPivot.Delete
            Application.DisplayAlerts = True
        End With
    End With
    
    Application.DisplayAlerts = False
    Worksheets("Pinpoint Reason Reference").Delete
    Worksheets("Raw Data").Delete
    Application.DisplayAlerts = True
    
    For Each shtApplicationPivot In Worksheets
        Call CreateCustomerFacingActiveReasons(Pvt, shtApplicationPivot)
        
        strApplicationName = shtApplicationPivot.Name
        Application.DisplayAlerts = False
        shtApplicationPivot.Delete
        Application.DisplayAlerts = True
        ActiveSheet.Name = strApplicationName
    Next shtApplicationPivot

'    strFileName = PrintToPDF
Application.ScreenUpdating = True
'    MsgBox "File is saved at:" & vbNewLine & strFileName
    Set Wbk = Nothing
    Set shtRawData = Nothing
End Sub

Function PrintToPDF() As String
    Dim oShell As Object
    Dim strFileName As String
    Dim strToday As String
    Dim strPrinterName As String
    
    strPrinterName = GetPrinterName("Microsoft Print to PDF")
    strToday = Format(Date, "yyyy-mm-dd")
    strFileName = GetDownloadsPath & Application.PathSeparator & "BPCE Risk Reasons per Activities " & strToday & ".pdf"
    ActiveWorkbook.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False, ActivePrinter:=strPrinterName, PrToFileName:="BPCE Risk Reasons per Activities.pdf"
    
    Set oShell = CreateObject("WScript.Shell")
    oShell.Run Chr(34) & strFileName & Chr(34)
    PrintToPDF = strFileName
End Function

Public Function GetPrinterName(ByVal PrinterName As String) As String

    'This works with Windows 2000 and up
    
    Dim Arr As Variant
    Dim Device As Variant
    Dim Devices As Variant
    Dim printer As String
    Dim RegObj As Object
    Dim RegValue As String
    Const HKEY_CURRENT_USER = &H80000001
    
    Set RegObj = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    RegObj.enumValues HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Devices", Devices, Arr
    
    For Each Device In Devices
        RegObj.getstringvalue HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Devices", Device, RegValue
        printer = Device & " on " & Split(RegValue, ",")(1)
        'If InStr(1, Printer, PrinterName, vbTextCompare) > 0 Then  'original code
        If StrComp(Device, PrinterName, vbTextCompare) = 0 Then
            GetPrinterName = printer
            Exit Function
        End If
    Next
      
End Function


