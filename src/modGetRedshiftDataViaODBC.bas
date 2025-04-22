Attribute VB_Name = "modGetRedshiftDataViaODBC"
Option Explicit

Private strQueryName As String
Private Const strGdId As String = "CEA4DAD6FA3F13714773DC23A60B5F898AFC4CA8E34EFB1E3615429F246EEFE6-28925943"
    
' Main procedure that orchestrates the entire process
Sub FetchAndDisplayRedshiftData()
    ' Create and execute the query
    CreateRedshiftQuery
    
    ' Add a new worksheet and populate it with query results
    DisplayQueryResults
End Sub

' Creates a query to fetch data from Redshift
Sub CreateRedshiftQuery()
    Dim sqlFormula As String
    
    Call Workbooks.Add(xlWBATWorksheet)
    
    strQueryName = RandomUniqueString
    sqlFormula = BuildSQLFormula()
    
    ' Add the query to the workbook
    ActiveWorkbook.Queries.Add name:=strQueryName, formula:=sqlFormula
End Sub

' Builds the SQL formula string for the Redshift query
Function BuildSQLFormula() As String
    Dim formula As String
    Dim part1 As String, part2 As String, part3 As String, part4 As String
    
    ' Break the long SQL string into multiple parts to avoid too many line continuations
    ' Adding proper spaces between SQL parts to maintain syntax
    part1 = "let" & Chr(13) & "" & Chr(10) & _
        "    Source = Odbc.Query(""dsn=Redshift_EU"", ""select distinct gd_id " & _
        "              , case " & _
        "                    when triggered_rule like '%cookie%' then CONCAT(cookie, ',') " & _
        "                    when triggered_rule like '%machine_id%' then CONCAT(machine_id, ',') " & _
        "                    when triggered_rule like '%gd_id%' then CONCAT(gd_id, ',') " & _
        "                    when triggered_rule like '%hw%' then CONCAT(hw_id, ',') " & _
        "                    when triggered_rule like '%advertising%' then CONCAT(advertising_id, ',') " & _
        "                    when triggered_rule like '%global%' then CONCAT(global_cookie, ',') end"
    
    part2 = " as attribute_value " & _
        "              , case " & _
        "                    when triggered_rule like '%cookie%' then 'cookie' " & _
        "                    when triggered_rule like '%machine_id%' then 'machine_id' " & _
        "                    when triggered_rule like '%gd_id%' then 'gd_id' " & _
        "                    when triggered_rule like '%hw%' then 'hw' " & _
        "                    when triggered_rule like '%advertising%' then 'advertising' " & _
        "                    when triggered_rule like '%global%' " & _
        "                        then 'global cookie' end"
    
    part3 = " as attribute_type " & _
        "              , case when triggered_rule like '%insight%' then 'White list' else 'Remove from block list' end as type " & _
        "from cagricole.policy_invocation_stats as pis " & _
        "         join cagricole.devices as d " & _
        "              on pis.session_id = d.session_id " & _
        "         left join cagricole.mobile_data as md " & _
        "               on md.session_id = d.session_id " & _
        "where gd_id in ('" & strGdId & "')"
    
    part4 = " and (score = 1000 or triggered_rule like '%known_fraudster%' or triggered_rule like '%peeking%') " & _
        "  and policy_category in ('L1_policies', 'L2_policies') " & _
        "  and d.created_at >= sysdate - 30 " & _
        " order by 4, 3;"")" & Chr(13) & "" & Chr(10) & _
        "in" & Chr(13) & "" & Chr(10) & _
        "    Source"
        
    ' Combine all parts without line continuation between them
    formula = part1
    formula = formula & part2
    formula = formula & part3
    formula = formula & part4
    
    BuildSQLFormula = formula
End Function

' Creates a new worksheet and displays the query results in a table
Sub DisplayQueryResults()
    Dim listObject As listObject
    Dim destinationRange As Range
    
    ' Set the destination for the query results
    Set destinationRange = Range("$A$1")
    
    ' Create a list object and query table
    Set listObject = ActiveSheet.ListObjects.Add( _
        SourceType:=0, _
        Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & strQueryName & ";Extended Properties=""""", _
        Destination:=destinationRange)
    
    ' Configure the query table properties
    ConfigureQueryTable listObject.queryTable, strQueryName
End Sub

' Configures the properties of a query table
Sub ConfigureQueryTable(queryTable As queryTable, queryName As String)
    With queryTable
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

