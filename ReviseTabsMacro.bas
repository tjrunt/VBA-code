Attribute VB_Name = "reviseTabs"
'''''  kill olds page tabs and create new & create a new PageName-Market tab named  "PageName-Market"  .... warning this currently does not delete the old PageName-Market renames it to "OLD-PageName-Market"

Sub reviseTabs_PageNameMarket()
    Call deleteTabs
    Call spinTabs
    Call renameTabs
    Call revisedMarketaddWS
    Worksheets("PageName-Market").Range("A1").Value = "WorkingPageID"
    Worksheets("PageName-Market").Range("B1").Value = "MarketCode"
    Call getMarketCodes
    Call moveTabs
End Sub

Private Sub renameTabs()
        Sheets("PageName-Market").Select
        myCurrentSheet = ActiveSheet.Name
        Sheets(myCurrentSheet).Name = "OLD-PageName-Market"

End Sub


'bulkReplace onlydoes col i and j of the selected tab
Sub bulkReplace()
        Dim count
        Dim mfindArry() As String
        Sheets("MasterData").Select
        myFindValue = InputBox("Paste in values to find separated by a space", "To Find", 1)
        mfindArry() = Split(myFindValue, " ")
        
        
        Dim replaceWrdArray() As String
        myreplaceValue = InputBox("Paste in values to repalce separated by a space,", "To Replace", 1)
        replaceWrdArray() = Split(myreplaceValue, " ")
        
       count = 0
   
   For Each Item In mfindArry
 
 
            
            Columns("I:J").Select
            Selection.Replace What:=Item, Replacement:=replaceWrdArray(count), LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=True

        
        count = count + 1
    Next

End Sub



' SORTS TABS BEing called in spintab function to activate to check  scripts
Private Sub mySortTabs()
 Dim wrksht As Worksheet
 Dim oListObj As ListObject
 
 Set wrksht = ActiveWorkbook.Worksheets(ActiveSheet.Name)
 Set oListObj = wrksht.ListObjects(1)
 
 
 ActiveWorkbook.Worksheets(ActiveSheet.Name).ListObjects(oListObj.Name).Sort.SortFields _
        .Clear
    ActiveWorkbook.Worksheets(ActiveSheet.Name).ListObjects(oListObj.Name).Sort.SortFields _
        .Add Key:=Range(oListObj.Name & "[[#All],[MarketName]]"), SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(ActiveSheet.Name).ListObjects(oListObj.Name).Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets(ActiveSheet.Name).ListObjects(oListObj.Name).Sort.SortFields _
        .Clear
    ActiveWorkbook.Worksheets(ActiveSheet.Name).ListObjects(oListObj.Name).Sort.SortFields _
        .Add Key:=Range(oListObj.Name & "[[#All],[MerchArea]]"), SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(ActiveSheet.Name).ListObjects(oListObj.Name).Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

' MsgBox oListObj.Name
End Sub

Private Sub deleteTabs()
    Dim pageNames
    Dim WrdArray() As String
    Application.DisplayAlerts = False
    
For i = 1 To Sheets.count

        currentPageName = Sheets(i).Name
        pageNames = pageNames & "|" & currentPageName

Next i
        
        WrdArray() = Split(pageNames, "|")

    For Each Item In WrdArray
      If IsNumeric(Left(Item, 1)) Then
            Sheets(Item).Select
    ActiveWindow.SelectedSheets.Delete
End If
   Next
   

        
    Application.DisplayAlerts = True
End Sub
Private Sub moveTabs()
    Worksheets("PivotTable").Move Before:=Worksheets(1)

    Worksheets("OLD-PageName-Market").Move Before:=Worksheets(2)
    Worksheets("PageName-Market").Move Before:=Worksheets(3)
End Sub


Private Sub spinTabs()
    Dim comH
    Dim startRow
    Dim PageName
    Dim myCurrentSheet
    PageName = "PivotTable"
    colA = "A"
    colH = "H"
    startRow = 2
    myCell = colH & startRow
    'Sheets("PivotTable").Select

    
    'todo skip grand total "Grand Total"
    ' todo move piviot  and other tabs to front
    '
    Do While Len(Worksheets(PageName).Range(myCell).Value) > 0 And Worksheets(PageName).Range(colA & startRow).Value <> "Grand Total"
        Sheets("PivotTable").Select
        Range(myCell).Select
        Selection.ShowDetail = True
        myCurrentSheet = ActiveSheet.Name
        Sheets(myCurrentSheet).Name = Worksheets(myCurrentSheet).Range("I2").Value
        
        
        'activate if need I was useing to check markets for testing against old MVRS
        'Call mySortTabs
        
        startRow = startRow + 1
        myCell = colH & startRow
Loop


End Sub

Private Sub getMarketCodes()
    Dim currentPageName As String
    Dim temp As String
    Dim RevisedPageNameMarket As String
    myVal = ""
    temp = ""
    pipe = "|"

    colC = "C"
    startRow = 2
    myCell = colC & startRow
    
    RevisedPageNameMarket = "PageName-Market"
    revisedStartRow = 2
    revisedColA = "A"
    revisedColB = "B"
    
    For i = 1 To Sheets.count
        currentPageName = Sheets(i).Name


  Do While Len(Worksheets(currentPageName).Range(myCell).Value) > 0

        temp = Worksheets(currentPageName).Range(myCell).Value
        temp = temp & pipe
        myVal = myVal & temp
        startRow = startRow + 1
        myCell = colC & startRow
        


    Loop


If IsNumeric(Left(currentPageName, 1)) Then


    Worksheets(RevisedPageNameMarket).Range(revisedColA & revisedStartRow).Value = currentPageName
    Worksheets(RevisedPageNameMarket).Range(revisedColB & revisedStartRow).Value = Left(myVal, Len(myVal) - 1)
End If
    'reset values
    myVal = ""
    temp = ""
    startRow = 2
    myCell = colC & startRow
    revisedStartRow = revisedStartRow + 1

Next i


End Sub

Private Sub revisedMarketaddWS()
    Dim ws As Worksheet
    Dim x  As String
  
     Set ws = Sheets.Add(after:=Sheets(Worksheets.count))
     ws.Name = "PageName-Market"
    

 End Sub


