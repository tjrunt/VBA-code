Attribute VB_Name = "piviotformat"
Private Sub oneToTwenty()


PageName = "PivotTable"
sourceCellNum = 2
Emycell = "E" & sourceCellNum
Dim newFolio As String
Dim myNextLetter As String

Do While Len(Worksheets(PageName).Range(Emycell).Value) > 0
        
    valEmycell = Worksheets(PageName).Range(Emycell).Value
    nextcell = sourceCellNum + 1
    valnextEmycell = Worksheets(PageName).Range("E" & nextcell).Value
    
  If IsNumeric(Right(valEmycell, 1)) And IsNumeric(Right(valnextEmycell, 1)) And InStr(valEmycell, "-") = 0 Then
        
        Worksheets(PageName).Range(Emycell).Value = Worksheets(PageName).Range(Emycell).Value & " (1-20)"
    ElseIf IsNumeric(Right(valEmycell, 1)) And valnextEmycell = "" And InStr(valEmycell, "-") = 0 Then
        Worksheets(PageName).Range(Emycell).Value = Worksheets(PageName).Range(Emycell).Value & " (1-20)"
  End If
  
  sourceCellNum = sourceCellNum + 1

Emycell = "E" & sourceCellNum

  Loop
    
End Sub


 

Sub piviot_format()
Attribute piviot_format.VB_Description = "formats the piviot on mvrs"
Attribute piviot_format.VB_ProcData.VB_Invoke_Func = "p\n14"
' IS GOOD
' piviot_format Macro
' formats the piviot on mvrs
'

'

Sheets("PivotTable").Select
    ActiveSheet.PivotTables("PivotTable2").PivotFields("PageSequenceID").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotTable2").PivotFields("WorkingPageID").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotTable2").PivotFields("PageID").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("PageName").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("PageFolio").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("MerchArea").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("MarketName").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
 ActiveSheet.PivotTables("PivotTable2").RowAxisLayout xlTabularRow
 ActiveSheet.PivotTables("PivotTable2").PivotFields("MerchArea").ShowDetail = False
 ActiveSheet.PivotTables("PivotTable2").PivotFields("PageFolio").ShowDetail = False
 
     Call oneToTwenty
 
 Columns("A:A").EntireColumn.AutoFit
    Columns("B:B").EntireColumn.AutoFit
    Columns("C:C").EntireColumn.AutoFit
    Columns("D:D").EntireColumn.AutoFit
    Columns("E:E").EntireColumn.AutoFit
    Columns("F:F").EntireColumn.AutoFit
    Columns("G:G").EntireColumn.AutoFit
    Columns("H:H").EntireColumn.AutoFit
    

End Sub







