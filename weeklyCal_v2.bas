Function dhFirstDayInWeek(Optional dtmDate As Date = 0) As Date
    ' Returns the first day in the week specified
    ' by the date in dtmDate.
    ' Uses localized settings for the first day of the week.
    If dtmDate = 0 Then
        ' Did the caller pass in a date? If not, use
        ' the current date.
        dtmDate = Date
    End If
    dhFirstDayInWeek = dtmDate - Weekday(dtmDate, _
     vbUseSystem) + 1
End Function
'''''''
Function dhLastDayInWeek(Optional dtmDate As Date = 0) As Date
    ' Returns the last day in the week specified by
    ' the date in dtmDate.
    ' Uses localized settings for the first day of the week.
    If dtmDate = 0 Then
        ' Did the caller pass in a date? If not, use
        ' the current date.
        dtmDate = Date
    End If
    dhLastDayInWeek = dtmDate - Weekday(dtmDate, vbUseSystem) + 7
End Function

Function getStageStartCol(currentSheet)


'todo pass these values as variables
myScrollColumn = ActiveWindow.ScrollColumn - 1 '30 AD
mySplitColumn = ActiveWindow.SplitColumn '5







If mySplitColumn = 0 Then
    getStageStartCol = 1
Else



 Do While mySplitColumn <= myScrollColumn

        If (Sheets(currentSheet).Columns(myScrollColumn).EntireColumn.Hidden = False) Then
           
            getStageStartCol = myScrollColumn - mySplitColumn + 1
            Exit Do
        End If
            
            
    myScrollColumn = myScrollColumn - 1
Loop

End If






End Function



Sub FormatWeeklyCal()

    Columns("A:H").Select
    Range("I1").Activate
    Selection.ColumnWidth = 41.67
    Range("A1:H2").Select
    With Selection
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlBottom: .WrapText = False: .Orientation = 0: .IndentLevel = 0: .ShrinkToFit = False: .ReadingOrder = xlContext:  .MergeCells = False
    End With
    Range("A1:A2").Select
    With Selection
    .HorizontalAlignment = xlCenter: .VerticalAlignment = xlBottom: .WrapText = False: .Orientation = 0: .IndentLevel = 0: .ShrinkToFit = False:  .ReadingOrder = xlContext:  .MergeCells = False
    End With
    Columns("A:H").Select
    With Selection
        .HorizontalAlignment = xlGeneral:        .VerticalAlignment = xlTop: .Orientation = 0:        .IndentLevel = 0:    .ShrinkToFit = False:      .ReadingOrder = xlContext:    .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlLeft:   .VerticalAlignment = xlTop:   .Orientation = 0:    .IndentLevel = 0:  .ShrinkToFit = False:     .ReadingOrder = xlContext:   .MergeCells = False:
    End With


    Selection.Replace What:="   ", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False
   
   
    Selection.Replace What:="  ", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False
    
    
    
       Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A:$H"), , xlYes).Name = _
        "Table1"
    Columns("A:H").Select
    Selection.ColumnWidth = 41.67
    ActiveSheet.ListObjects("Table1").TableStyle = "TableStyleLight16"
      Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    
    End Sub
    
 
Sub runWeeklySchedule()

''''Start Optimize code
Dim CalcState As Long
Dim EventState As Boolean
Dim PageBreakState As Boolean

Application.ScreenUpdating = False

EventState = Application.EnableEvents
Application.EnableEvents = False

CalcState = Application.Calculation
Application.Calculation = xlCalculationManual

PageBreakState = ActiveSheet.DisplayPageBreaks
ActiveSheet.DisplayPageBreaks = False
''''''End


 

Call get_weeklyMileStones


''deactivate optimize code
ActiveSheet.DisplayPageBreaks = PageBreakState
Application.Calculation = CalcState
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = EventState
Application.ScreenUpdating = True
'deactivate optimize code End

MsgBox "Complete"

End Sub



Private Sub get_weeklyMileStones()

startDate = dhFirstDayInWeek()
endDate = dhLastDayInWeek()
 


 Dim sun As Date
 Dim mon As Date
 Dim tues As Date
 Dim wed As Date
 Dim thurs As Date
 Dim fri As Date
Dim sat As Date

Dim destSunCellRowNum As Integer
Dim destMonCellRowNum As Integer
Dim destTuesCellRowNum As Integer
Dim destWedCellRowNum As Integer
Dim destThursCellRowNum As Integer
Dim destFriCellRowNum As Integer
Dim destSatCellRowNum As Integer
Dim destMultiCellRowNum As Integer

Dim mySheetName As String
Dim SheetName As String


Dim SplitColumnCount As Integer
Dim ScrollColumnNumber As Integer


Dim myRow   As Integer
Dim myColumn  As Integer



    
    Dim str_sun As String
    Dim str_mon As String
    Dim str_tue As String
    Dim str_wed As String
    Dim str_thu As String
    Dim str_fri As String
    Dim str_sat As String


    
    If (InStr(1, Application.OperatingSystem, "mac") <> 0) Then
     trimVAl = 5
    Else
        trimVAl = 3
 End If
     
    
    

    
    
    
    

'make Sheet for weekly cal Start
    sun = startDate
    mon = sun + 1
    tues = sun + 2
    wed = sun + 3
    thurs = sun + 4
    fri = sun + 5
    sat = endDate
 
 
 
 
     str_sun = sun
    str_mon = mon
    str_tue = tues
    str_wed = wed
    str_thu = thurs
    str_fri = fri
    str_sat = sat
    
    

    
    str_sun = Left(str_sun, Len(str_sun) - trimVAl)
    str_mon = Left(str_mon, Len(str_mon) - trimVAl)
    str_tue = Left(str_tue, Len(str_tue) - trimVAl)
    str_wed = Left(str_wed, Len(str_wed) - trimVAl)
    str_thu = Left(str_thu, Len(str_thu) - trimVAl)
    str_fri = Left(str_fri, Len(str_fri) - trimVAl)
    str_sat = Left(str_sat, Len(str_sat) - trimVAl)
    
    
 
 
 
 
 
 
 
 
 
    
    
'''''
''''''
    mySheetName = "week-of-" & CStr(startDate)
    mySheetName = Replace(mySheetName, "/", "-")
    Sheets.Add.Name = mySheetName
    
   
 
    Worksheets(mySheetName).Range("A1").Value = "Sunday" & vbNewLine & sun
    Worksheets(mySheetName).Range("B1").Value = "Monday" & vbNewLine & mon
    Worksheets(mySheetName).Range("C1").Value = "Tuesday" & vbNewLine & tues
    Worksheets(mySheetName).Range("D1").Value = "Wednesday" & vbNewLine & wed
    Worksheets(mySheetName).Range("E1").Value = "Thursday" & vbNewLine & thurs
    Worksheets(mySheetName).Range("F1").Value = "Friday" & vbNewLine & fri
    Worksheets(mySheetName).Range("G1").Value = "Saturday" & vbNewLine & sat
    Worksheets(mySheetName).Range("H1").Value = "Multiday Day Tasks"






'''''
''''''
   destSunCellRowNum = 3
    destMonCellRowNum = 3
    destTuesCellRowNum = 3
    destWedCellRowNum = 3
    destThursCellRowNum = 3
    destFriCellRowNum = 3
    destSatCellRowNum = 3
    destMultiCellRowNum = 3
'make Sheet for weekly cal End

For Each ws In ActiveWorkbook.Worksheets
 
    
    SheetName = ws.Name
    Sheets(SheetName).Activate
    

    
  
    
    If InStr(1, SheetName, "week-of-") = 0 Then
     
     
     ' goto home and set scroll col and freeze count
      ''Cells.Select
      '''Selection.EntireColumn.Hidden = False
      ActiveWindow.ScrollColumn = 1
      '
      
      
      'todo update  getStage function and parameters nolonger need 0 logic
      'SplitColumnCount = ActiveWindow.SplitColumn 'number of SplitColumn that are  split
      'ScrollColumnNumber = ActiveWindow.ScrollColumn 'number of SplitColumn that are  split
      
      StageStartCol = getStageStartCol(SheetName)
      


       For Each cell In ws.UsedRange.Cells
        If (cell.Value <> vbNullString) Then
        'MsgBox cell.Value
            If (cell.Value >= sun) And (cell.Value <= sat) And (cell.Font.ColorIndex <> 6) Then
             'MsgBox " yes"
       
                            If (cell.Value = sun) Then

                                stagecell = getStage(StageStartCol, SheetName, cell.Row, cell.Column)
                                
                                
                                If (stagecell <> 0) Then
                                    Worksheets(mySheetName).Range("A" & destSunCellRowNum).Value = stagecell
                                    destSunCellRowNum = destSunCellRowNum + 1
                                End If

                 '
                ''mon start
                '
                                         
                            ElseIf (cell.Value = mon) Then

                            
                            
                                stagecell = getStage(StageStartCol, SheetName, cell.Row, cell.Column)
                                If (stagecell <> 0) Then
                                    Worksheets(mySheetName).Range("B" & destMonCellRowNum).Value = stagecell
                                    destMonCellRowNum = destMonCellRowNum + 1
                                End If
                          
                '
                'mon end
                '
                
                                 '
                ''tues start
                '
                                         
                            ElseIf (cell.Value = tues) Then

                            
                            
                                stagecell = getStage(StageStartCol, SheetName, cell.Row, cell.Column)
                                If (stagecell <> 0) Then
                                    Worksheets(mySheetName).Range("C" & destTuesCellRowNum).Value = stagecell
                                    destTuesCellRowNum = destTuesCellRowNum + 1
                                End If
                            
                '
                'tues end
                '
                
                
                
                                 '
                ''wed start
                '
                                         
                            ElseIf (cell.Value = wed) Then

                            
                            
                                stagecell = getStage(StageStartCol, SheetName, cell.Row, cell.Column)
                                If (stagecell <> 0) Then
                                    Worksheets(mySheetName).Range("D" & destWedCellRowNum).Value = stagecell
                                    destWedCellRowNum = destWedCellRowNum + 1
                                End If
                           
                '
                'wed end
                '
                
                
                                 '
                ''thurs start
                '
                                         
                            ElseIf (cell.Value = thurs) Then

                            
                            
                                stagecell = getStage(StageStartCol, SheetName, cell.Row, cell.Column)
                                If (stagecell <> 0) Then
                                    Worksheets(mySheetName).Range("E" & destThursCellRowNum).Value = stagecell
                                    destThursCellRowNum = destThursCellRowNum + 1
                                End If
                            
                '
                'thurs end
                '
                
                
                                 '
                ''fri start
                '
                                         
                            ElseIf (cell.Value = fri) Then

                            
                            
                                stagecell = getStage(StageStartCol, SheetName, cell.Row, cell.Column)
                                If (stagecell <> 0) Then
                                    Worksheets(mySheetName).Range("F" & destFriCellRowNum).Value = stagecell
                                    destFriCellRowNum = destFriCellRowNum + 1
                                End If
                           
                '
                'fri end
                '
                
                
                                 '
                ''sat start
                '
                                         
                            ElseIf (cell.Value = sat) Then

                            
                            
                                stagecell = getStage(StageStartCol, SheetName, cell.Row, cell.Column)
                                If (stagecell <> 0) Then
                                    Worksheets(mySheetName).Range("G" & destSatCellRowNum).Value = stagecell
                                    destSatCellRowNum = destSatCellRowNum + 1
                                End If
                   End If
                                
                                
                    'rreanges dates
                    
            Else
                    If (cell.Font.ColorIndex <> 6) And ((InStr(1, cell.Value, "-")) <> 0) Then
                         If ((InStr(1, cell.Value, str_mon)) <> 0) Or ((InStr(1, cell.Value, str_wed)) <> 0) Or ((InStr(1, cell.Value, str_thu)) <> 0) Or ((InStr(1, cell.Value, str_fri)) <> 0) Or ((InStr(1, cell.Value, str_tue)) <> 0) Or ((InStr(1, cell.Value, str_sat)) <> 0) Or ((InStr(1, cell.Value, str_sun)) <> 0) Then

                                    stagecell = getStage(StageStartCol, SheetName, cell.Row, cell.Column)
                                     If (stagecell <> 0) Then
                                         stagecell = stagecell & vbNewLine & " Task days active:" & cell.Value
                                        Worksheets(mySheetName).Range("H" & destMultiCellRowNum).Value = stagecell
                                        destMultiCellRowNum = destMultiCellRowNum + 1
                                        'MsgBox " in range if"
                                     End If
                        End If
                    End If
                                
                                
                                
            End If
                                
        End If
        

            
         Next cell
        'instring if end
        End If
    Next ws


Sheets(mySheetName).Select



Call FormatWeeklyCal

End Sub



'getStage info
Function getStage(StageStartCol, mySheet, myRow, myColumn)
Dim myCellVal As String
Dim stagArry(6) As String
Dim tempCount As Byte
Dim StageendCol As Integer






myCellVal = ""
stagArry(1) = "Wk: "
stagArry(2) = "Team: "
stagArry(3) = "Lago: "
stagArry(4) = "Action: "
stagArry(5) = ""

tempStageStartCol = StageStartCol
StageendCol = StageStartCol + 5

'todo Add statement to loop through the desired stage columns
'select range of next 5 coluns in specific row
' if its not a date build the data

tempCount = 1
On Error Resume Next
 Do While tempStageStartCol <= StageendCol
    If (IsDate(Worksheets(mySheet).Cells(myRow, tempStageStartCol).Value) <> True) Then

        myCellVal = myCellVal & " " & stagArry(tempCount) & Worksheets(mySheet).Cells(myRow, tempStageStartCol).Value & vbNewLine
        tempStageStartCol = tempStageStartCol + 1
        tempCount = tempCount + 1
    Else
        Exit Do
        
    End If
    
Loop



'checks for desired stages
If InStr(1, myCellVal, "FCC/IMC/Creative") <> 0 Or InStr(1, myCellVal, "Traffic/OPS") <> 0 Or InStr(1, myCellVal, "Layout Admin") <> 0 Or InStr(1, myCellVal, "Traffic/OPS") <> 0 Or InStr(1, myCellVal, "PAs") <> 0 _
    Or InStr(1, myCellVal, "Editors") <> 0 Or InStr(1, myCellVal, "Creative Director") <> 0 Or InStr(1, myCellVal, "Traffic/Layout/BA") <> 0 Or InStr(1, myCellVal, "Account/Merchants") <> 0 Or InStr(1, myCellVal, "Traffic") <> 0 _
    Or InStr(1, myCellVal, "PAs/Editors") <> 0 Or InStr(1, myCellVal, "QC PA/Business Analyst") <> 0 Or InStr(1, myCellVal, "Proof WAM (Wed)") <> 0 Then
        v4 = "AD Date: " & Cells(3, myColumn).Value & vbNewLine
        v5 = "Activity ID: " & Cells(6, myColumn).Value & vbNewLine
        v6 = "Ad Manager: " & Cells(10, myColumn).Value & vbNewLine
        
        'MsgBox "counter:  " & counter
        getStage = mySheet & vbNewLine & myCellVal & vbNewLine & v4 & v5 & v6


Else
    getStage = 0


End If



End Function




















