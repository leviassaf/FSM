Attribute VB_Name = "modCreateSavedTrxAmounts"
Option Explicit

Private Const FROM_EUROCENT_TO_EURO As Double = 100#
' --- Constants for alignment and cell formatting ---
Const DEFAULT_HORIZONTAL_ALIGNMENT = xlCenter    ' -4108
Const DEFAULT_VERTICAL_ALIGNMENT = xlBottom      ' -4107
Const DEFAULT_WRAP_TEXT = False
Const DEFAULT_ORIENTATION = 0
Const DEFAULT_ADD_INDENT = False
Const DEFAULT_INDENT_LEVEL = 0
Const DEFAULT_SHRINK_TO_FIT = False
Const DEFAULT_READING_ORDER = xlContext         ' -5002
Const DEFAULT_MERGE_CELLS = False

' --- Constants for border formatting ---
Const BORDER_LINE_STYLE = xlContinuous          ' 1
Const BORDER_COLOR_INDEX = 0
Const BORDER_TINT_AND_SHADE = 0
Const BORDER_WEIGHT = xlThin                    ' 2

' --- Constant for numeric constants in SpecialCells ---
Const NUMERIC_CONSTANTS = 1                     ' 1 means numeric cells

Sub TransformCSVsToMoneySaved()
    Dim folderPath As String
    Dim mainQueryName As String, sampleQueryName As String
    Dim parameterQueryName As String, transformSampleQueryName As String
    Dim transformFileQueryName As String
    
    Dim mainQueryFormula As String, sampleQueryFormula As String
    Dim parameterQueryFormula As String, transformSampleQueryFormula As String
    Dim transformFileQueryFormula As String
    
    ' Prompt user to select the folder containing CSV files.
    folderPath = GetFolderPath()
    If folderPath = "" Then
        MsgBox "No folder selected. Exiting.", vbExclamation
        Exit Sub
    End If
    
    ' Check if the folder contains any CSV files.
    If Not FolderContainsCSV(folderPath) Then
        MsgBox "The selected folder does not contain any CSV files.", vbExclamation
        Exit Sub
    End If
    ' Create a new workbook with exactly one worksheet
    Call Workbooks.Add(xlWBATWorksheet)
    
    ' Define query names (change these if desired).
    mainQueryName = "ReportQuery"
    sampleQueryName = "SampleFile"
    parameterQueryName = "ParameterQuery"
    transformSampleQueryName = "TransformSampleFile"
    transformFileQueryName = "TransformFile"
    
    ' Build the M code for each query.
    mainQueryFormula = BuildMainQueryFormula(folderPath, sampleQueryName)
    sampleQueryFormula = BuildSampleQueryFormula(folderPath)
    parameterQueryFormula = BuildParameterQueryFormula(sampleQueryName)
    transformSampleQueryFormula = BuildTransformSampleQueryFormula(parameterQueryName)
    transformFileQueryFormula = BuildTransformFileQueryFormula(parameterQueryName)
    
    ' Add each query to the workbook.
    AddWorkbookQuery mainQueryName, mainQueryFormula
    AddWorkbookQuery sampleQueryName, sampleQueryFormula
    AddWorkbookQuery parameterQueryName, parameterQueryFormula
    AddWorkbookQuery transformSampleQueryName, transformSampleQueryFormula
    AddWorkbookQuery transformFileQueryName, transformFileQueryFormula
    
    ' Create a new sheet to display the final report.
    CreateReportWorksheet mainQueryName
    Call FormatReport
    Call DeleteAllSheetsExceptActive
    Call PageSetupForPrint
End Sub

Private Sub FormatDataAsCurrency()
    Dim rngData As Range
    
    Set rngData = ActiveSheet.Range("A1").CurrentRegion
    
End Sub

' ==============================
' =         UTILITIES         =
' ==============================

' Prompts the user to pick a folder.
Private Function GetFolderPath() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Select Folder Containing CSV Files"
        If .Show = -1 Then
            GetFolderPath = .SelectedItems(1)
        Else
            GetFolderPath = ""
        End If
    End With
End Function

' Checks if the specified folder contains any CSV files.
Private Function FolderContainsCSV(ByVal folderPath As String) As Boolean
    Dim fileName As String
    fileName = Dir(folderPath & "\*.csv")
    If fileName = "" Then
        FolderContainsCSV = False
    Else
        FolderContainsCSV = True
    End If
End Function

' Safely adds a query to the active workbook.
Private Sub AddWorkbookQuery(queryName As String, queryFormula As String)
    On Error Resume Next
    ActiveWorkbook.Queries.Add name:=queryName, formula:=queryFormula
    On Error GoTo 0
End Sub

' Creates a new worksheet and loads the main query result into a table.
Private Sub CreateReportWorksheet(ByVal mainQueryName As String)
    Dim ws As Worksheet
    Dim displayName As String
    
    ' Create a new worksheet.
    Set ws = ActiveWorkbook.Worksheets.Add
    ws.name = "CSV Report"
    
    displayName = "_" & Replace(mainQueryName, " ", "_") & "_"
    
    With ws.ListObjects.Add( _
            SourceType:=0, _
            Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""" & mainQueryName & """;Extended Properties=""""", _
            Destination:=ws.Range("$A$1")).queryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & mainQueryName & "]")
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
        .listObject.displayName = displayName
        .Refresh BackgroundQuery:=False
    End With
    ws.ListObjects(1).Unlist
End Sub

' ==============================
' =       QUERY BUILDERS      =
' ==============================

' Builds the main Power Query that combines and pivots all CSV data.
Private Function BuildMainQueryFormula(ByVal folderPath As String, ByVal sampleQueryName As String) As String
    Dim lines() As String
    ReDim lines(1 To 31)
    
    lines(1) = "let"
    lines(2) = "    Divisor = " & FROM_EUROCENT_TO_EURO & ","
    lines(3) = "    Source = Folder.Files(""" & folderPath & """),"
    lines(4) = "    FilteredHiddenFiles = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true),"
    lines(5) = "    InvokeCustomFunction = Table.AddColumn(FilteredHiddenFiles, ""Transform File"", each TransformFile([Content])),"
    lines(6) = "    RenamedColumns = Table.RenameColumns(InvokeCustomFunction, {{""Name"", ""Source.Name""}}),"
    lines(7) = "    RemovedOtherColumns = Table.SelectColumns(RenamedColumns, {""Source.Name"", ""Transform File""}),"
    lines(8) = "    ExpandedTableColumn = Table.ExpandTableColumn(RemovedOtherColumns, ""Transform File"", "
    lines(9) = "        Table.ColumnNames(TransformFile(#""" & sampleQueryName & """))),"
    lines(10) = "    ChangedType = Table.TransformColumnTypes(ExpandedTableColumn, {"
    lines(11) = "        {""Source.Name"", type text},"
    lines(12) = "        {""Months in transaction_date"", type date},"
    lines(13) = "        {""Money Saved"", type number},"
    lines(14) = "        {""Money Loss"", type any},"
    lines(15) = "        {""%Money Saved"", Int64.Type}"
    lines(16) = "    }),"
    lines(17) = "    RemovedColumns = Table.RemoveColumns(ChangedType, {""Money Loss"", ""%Money Saved""}),"
    lines(18) = "    ReplacedValue = Table.ReplaceValue(RemovedColumns,"".csv"","""",Replacer.ReplaceText,{""Source.Name""}),"
    lines(19) = "    ReplacedValue1 = Table.ReplaceValue(ReplacedValue,""_"","" "",Replacer.ReplaceText,{""Source.Name""}),"
    lines(20) = "    CapitalizedEachWord = Table.TransformColumns(ReplacedValue1, {{""Source.Name"", Text.Proper, type text}}),"
    lines(21) = "    PivotedColumn = Table.Pivot(CapitalizedEachWord, "
    lines(22) = "        List.Distinct(CapitalizedEachWord[Source.Name]), ""Source.Name"", ""Money Saved"", List.Sum),"
    lines(23) = "    AddedCustom = Table.AddColumn(PivotedColumn, ""Year-Month"", each "
    lines(24) = "        Number.ToText(Date.Year([Months in transaction_date])) & ""-"" & "
    lines(25) = "        Text.Start(Date.MonthName([Months in transaction_date]), 3)),"
    lines(26) = "    RemovedColumns1 = Table.RemoveColumns(AddedCustom, {""Months in transaction_date""}),"
    lines(27) = "    ReorderedColumns = Table.ReorderColumns(RemovedColumns1, {""Year-Month"", ""Account Takeover"", ""First Party"", ""Mule Account"", ""Remote Access Tool"", ""Social Engineering"", ""Stolen Device""}),"

    lines(28) = "    ReplacedValue2 = Table.ReplaceValue(#""Pivoted Column"",null,0,Replacer.ReplaceValue,{""Account Takeover"", ""First Party"", ""Mule Account"", ""Remote Access Tool"", ""Social Engineering"", ""Stolen Device""}),"
    lines(29) = "    AddedCustom1 = Table.AddColumn(#""Replaced Value2"", ""Grand Total"", each [Account Takeover]+[First Party]+[Mule Account]+[Remote Access Tool]+[Social Engineering]+[Stolen Device]),"
    
    lines(30) = "    DividedColumns = Table.TransformColumns(ReorderedColumns, {"
    lines(31) = "        {""Account Takeover"", each Number.Round(_ / Divisor, 0), type number}, {""First Party"", each Number.Round(_ / Divisor, 0), type number}, {""Mule Account"", each Number.Round(_ / Divisor, 0), type number}, {""Remote Access Tool"", each Number.Round(_ / Divisor, 0), type number}, {""Social Engineering"", each Number.Round(_ / Divisor, 0), type number}, {""Stolen Device"", each Number.Round(_ / Divisor, 0), type number} }) in DividedColumns"
    
    BuildMainQueryFormula = Join(lines, vbCrLf)
End Function

' Builds the sample file query formula using the folder path.
Private Function BuildSampleQueryFormula(ByVal folderPath As String) As String
    Dim lines() As String
    ReDim lines(1 To 5)
    
    lines(1) = "let"
    lines(2) = "    Source = Folder.Files(""" & folderPath & """),"
    lines(3) = "    Navigation1 = Source{0}[Content]"
    lines(4) = "in"
    lines(5) = "    Navigation1"
    
    BuildSampleQueryFormula = Join(lines, vbCrLf)
End Function

' Builds the parameter query formula referencing the sample file query.
Private Function BuildParameterQueryFormula(ByVal sampleQueryName As String) As String
    Dim lines() As String
    ReDim lines(1 To 5)
    
    ' We don't need line breaks for meta definitions, but we'll keep them minimal:
    lines(1) = "#""" & sampleQueryName & """ meta ["
    lines(2) = "    IsParameterQuery=true, "
    lines(3) = "    BinaryIdentifier=#""" & sampleQueryName & """, "
    lines(4) = "    Type=""Binary"", IsParameterQueryRequired=true]"
    lines(5) = ""
    
    BuildParameterQueryFormula = Join(lines, "")
End Function

' Builds the transform sample file query formula.
Private Function BuildTransformSampleQueryFormula(ByVal parameterQueryName As String) As String
    Dim lines() As String
    ReDim lines(1 To 5)
    
    lines(1) = "let"
    lines(2) = "    Source = Csv.Document(" & parameterQueryName & ", [Delimiter="","", Columns=4, Encoding=65001, QuoteStyle=QuoteStyle.None]),"
    lines(3) = "    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true])"
    lines(4) = "in"
    lines(5) = "    PromotedHeaders"
    
    BuildTransformSampleQueryFormula = Join(lines, vbCrLf)
End Function

' Builds the transform file query formula as a function.
Private Function BuildTransformFileQueryFormula(ByVal parameterQueryName As String) As String
    Dim lines() As String
    ReDim lines(1 To 7)
    
    lines(1) = "let"
    lines(2) = "    Source = (" & parameterQueryName & ") => let"
    lines(3) = "        Source = Csv.Document(" & parameterQueryName & ", [Delimiter="","", Columns=4, Encoding=65001, QuoteStyle=QuoteStyle.None]),"
    lines(4) = "        PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true])"
    lines(5) = "    in"
    lines(6) = "        PromotedHeaders"
    lines(7) = "in"
    
    BuildTransformFileQueryFormula = Join(lines, vbCrLf) & vbCrLf & "    Source"
End Function

' ----------------------------------------------------
' Applies alignment and cell formatting properties
' ----------------------------------------------------
Private Sub FormatAlignment(rng As Range)
    With rng
        .HorizontalAlignment = DEFAULT_HORIZONTAL_ALIGNMENT
        .VerticalAlignment = DEFAULT_VERTICAL_ALIGNMENT
        .WrapText = DEFAULT_WRAP_TEXT
        .Orientation = DEFAULT_ORIENTATION
        .AddIndent = DEFAULT_ADD_INDENT
        .IndentLevel = DEFAULT_INDENT_LEVEL
        .ShrinkToFit = DEFAULT_SHRINK_TO_FIT
        .ReadingOrder = DEFAULT_READING_ORDER
        .MergeCells = DEFAULT_MERGE_CELLS
    End With
End Sub

' ----------------------------------------------------
' Removes diagonal borders from the specified range
' ----------------------------------------------------
Private Sub RemoveDiagonalBorders(rng As Range)
    With rng.Borders(xlDiagonalDown)
        .LineStyle = xlNone
    End With
    With rng.Borders(xlDiagonalUp)
        .LineStyle = xlNone
    End With
End Sub

' ----------------------------------------------------
' Applies border formatting to all specified edges
' ----------------------------------------------------
Private Sub FormatBorders(rng As Range)
    Dim borderIndices As Variant
    Dim i As Long
    
    borderIndices = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
    
    For i = LBound(borderIndices) To UBound(borderIndices)
        With rng.Borders(borderIndices(i))
            .LineStyle = BORDER_LINE_STYLE
            .ColorIndex = BORDER_COLOR_INDEX
            .TintAndShade = BORDER_TINT_AND_SHADE
            .Weight = BORDER_WEIGHT
        End With
    Next i
End Sub

Sub FormatReport()
    Dim reportRange As Range
    ' Identify the "report" range based on the ActiveCell's current region
    Set reportRange = ActiveCell.CurrentRegion
    
    ' Apply all formatting steps
    FormatAlignment reportRange
    RemoveDiagonalBorders reportRange
    FormatBorders reportRange
    
    ' NOTE: If you need a Euro symbol, try "#,##0 [$€-fr-FR]" or remove it if it causes issues
    ApplyNumberFormatToConstants reportRange, "#,##0 " & ChrW(8364)
End Sub

' Applies a specified number format to numeric constants
' ----------------------------------------------------
Private Sub ApplyNumberFormatToConstants(rng As Range, formatString As String)
    Dim constantCells As Range
    
    On Error Resume Next  ' In case there are no numeric constant cells
    Set constantCells = rng.SpecialCells(xlCellTypeConstants, NUMERIC_CONSTANTS)
    On Error GoTo 0
    
    If Not constantCells Is Nothing Then
        constantCells.NumberFormat = formatString
    End If
End Sub

Sub PageSetupForPrint()
    With ActiveSheet.PageSetup
        .PrintArea = ActiveSheet.UsedRange.Address
        .CenterHeader = _
        "&""-,Bold""&18LCL EUR Amount Saved" & Chr(10) & "January 2025 - March 2025"
        .LeftMargin = Application.InchesToPoints(0)
        .RightMargin = Application.InchesToPoints(0)
        .TopMargin = Application.InchesToPoints(1.34)
        .CenterHorizontally = True
        .Orientation = xlLandscape
        .Zoom = 130
    End With
End Sub




