Attribute VB_Name = "WeeklyCalendar"
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






Private Sub get_weeklyMileStones(startDate, endDate)

 
    sun = startDate
    mon = sun + 1
    tues = sun + 2
    wed = sun + 3
    thurs = sun + 4
    fri = sun + 5
    sat = endDate
    

'''''
''''''
    mySheetName = "week-of-" & CStr(startDate)
    mySheetName = Replace(mySheetName, "/", "-")
    Sheets.Add.Name = mySheetName
    
   
 
    Worksheets(mySheetName).Range("A1").Value = "Sunday"
    Worksheets(mySheetName).Range("B1").Value = "Monday"
    Worksheets(mySheetName).Range("C1").Value = "Tuesday"
    Worksheets(mySheetName).Range("D1").Value = "Wednesday"
    Worksheets(mySheetName).Range("E1").Value = "Thursday"
    Worksheets(mySheetName).Range("F1").Value = "Friday"
    Worksheets(mySheetName).Range("G1").Value = "Saturday"
    Worksheets(mySheetName).Range("A2").Value = sun
    Worksheets(mySheetName).Range("B2").Value = mon
    Worksheets(mySheetName).Range("C2").Value = tues
    Worksheets(mySheetName).Range("D2").Value = wed
    Worksheets(mySheetName).Range("E2").Value = thurs
    Worksheets(mySheetName).Range("F2").Value = fri
    Worksheets(mySheetName).Range("G2").Value = sat
    
'''''
''''''

For Each ws In ActiveWorkbook.Worksheets
    destSunCellRowNum = 3
    destMonCellRowNum = 3
    destTuesCellRowNum = 3
    destWedCellRowNum = 3
    destThursCellRowNum = 3
    destFriCellRowNum = 3
    destSatCellRowNum = 3
    
    sheetName = ws.Name
    Sheets(sheetName).Activate
        For Each Cell In ws.UsedRange.Cells
            myRow = Cell.Row
            myColumn = Cell.Column
            'MsgBox "in for loop "
            If (Cell.Value = sun) Then
                'lRow = Worksheets(mySheetName).Cells(Rows.count, 1).End(xlUp).Row
                         
            ElseIf (Cell.Value = mon) Then
                'lRow = Worksheets(mySheetName).Cells(Rows.count, 2).End(xlUp).Row
                
                v1 = Cells(myRow, "B").Value
                v2 = Cells(myRow, "C").Value
                v3 = Cells(myRow, "D").Value
                v4 = Cells(3, myColumn).Value
                v5 = Cells(6, myColumn).Value
                
                 myCell = sheetName & " " & v1 & " " & v2 & " " & v3 & " AD Dates: " & v4 & " Activity ID: " & v5
                 
                 'MsgBox myCell

                 Worksheets(mySheetName).Range("B" & destMonCellRowNum).Value = myCell
                 destMonCellRowNum = destMonCellRowNum + 1
                 
            ElseIf (Cell.Value = tues) Then
                'lRow = Cells(Rows.count, 3).End(xlUp).Row + 1
                               
                v1 = Cells(myRow, "B").Value
                v2 = Cells(myRow, "C").Value
                v3 = Cells(myRow, "D").Value
                v4 = Cells(3, myColumn).Value
                v5 = Cells(6, myColumn).Value
                
                 myCell = sheetName & " " & v1 & " " & v2 & " " & v3 & " AD Dates: " & v4 & " Activity ID: " & v5
                 

                 Worksheets(mySheetName).Range("C" & destTuesCellRowNum).Value = myCell
                 destTuesCellRowNum = destTuesCellRowNum + 1
                 
            ElseIf (Cell.Value = wed) Then
                'lRow = Cells(Rows.count, 4).End(xlUp).Row + 1
                               
                v1 = Cells(myRow, "B").Value
                v2 = Cells(myRow, "C").Value
                v3 = Cells(myRow, "D").Value
                v4 = Cells(3, myColumn).Value
                v5 = Cells(6, myColumn).Value
                
                 myCell = sheetName & " " & v1 & " " & v2 & " " & v3 & " AD Dates: " & v4 & " Activity ID: " & v5
                 

                 Worksheets(mySheetName).Range("D" & destWedCellRowNum).Value = myCell
                 destWedCellRowNum = destWedCellRowNum + 1
                 
            ElseIf (Cell.Value = thurs) Then
                'lRow = Cells(Rows.count, 5).End(xlUp).Row + 1
                               
                v1 = Cells(myRow, "B").Value
                v2 = Cells(myRow, "C").Value
                v3 = Cells(myRow, "D").Value
                v4 = Cells(3, myColumn).Value
                v5 = Cells(6, myColumn).Value
                
                 myCell = sheetName & " " & v1 & " " & v2 & " " & v3 & " AD Dates: " & v4 & " Activity ID: " & v5

                 Worksheets(mySheetName).Range("E" & destThursCellRowNum).Value = myCell
                 destThursCellRowNum = destThursCellRowNum + 1
                 
                 
            ElseIf (Cell.Value = fri) Then
                'lRow = Cells(Rows.count, 6).End(xlUp).Row + 1
                               
                v1 = Cells(myRow, "B").Value
                v2 = Cells(myRow, "C").Value
                v3 = Cells(myRow, "D").Value
                v4 = Cells(3, myColumn).Value
                v5 = Cells(6, myColumn).Value
                
                 myCell = sheetName & " " & v1 & " " & v2 & " " & v3 & " AD Dates: " & v4 & " Activity ID: " & v5
                 

                 Worksheets(mySheetName).Range("F" & destFriCellRowNum).Value = myCell
                 destFriCellRowNum = destFriCellRowNum + 1
                 
                 
            ElseIf (Cell.Value = sat) Then
                'lRow = Cells(Rows.count, 7).End(xlUp).Row + 1
                               
                v1 = Cells(myRow, "B").Value
                v2 = Cells(myRow, "C").Value
                v3 = Cells(myRow, "D").Value
                v4 = Cells(3, myColumn).Value
                v5 = Cells(6, myColumn).Value
                
                 myCell = sheetName & " " & v1 & " " & v2 & " " & v3 & " AD Dates: " & v4 & " Activity ID: " & v5
                 

                 Worksheets(mySheetName).Range("B" & destMonCellRowNum).Value = myCell
                 destMonCellRowNum = destMonCellRowNum + 1
                 
                 
            End If
            
         Next Cell
    
    Next ws

    






End Sub
Sub ZZZZZZZZ_firstEmptyRow()
    Dim lRow As Long
    lRow = Cells(Rows.count, 1).End(xlUp).Row
    'MsgBox lRow
    lRow = Cells(Rows.count, 1).End(xlUp).Row + 1
    'MsgBox lRow
End Sub

'https://www.excelcampus.com/vba/find-last-row-column-cell/
 Sub zzRange_End_Method()
'Finds the last non-blank cell in a single row or column

Dim lRow As Long
Dim lCol As Long
    
    'Find the last non-blank cell in column A(1)
    lRow = Cells(Rows.count, 7).End(xlUp).Row
    fEmptyRow = lRow + 1
    'Find the last non-blank cell in row 1
    lCol = Cells(1, Columns.count).End(xlToLeft).Column
    
    MsgBox "Last Row: " & fEmptyRow & vbNewLine & _
            "Last Column: " & lCol
            
            
    'Range("B3").Value = 2
    
End Sub


Sub ZZ_Last_Row()
x = firstEmptyRow()
MsgBox x
End Sub


Private Sub xxxmynewsheet(startDay)
'
' zzzzznewsheet Macro
'

'
    mySheetName = "week-of-" & CStr(startDay)
    mySheetName = Replace(mySheetName, "/", "-")
    Sheets.Add.Name = mySheetName
    
   
 
    Worksheets(mySheetName).Range("A1").Value = "Sunday"
    Worksheets(mySheetName).Range("B1").Value = "Monday"
    Worksheets(mySheetName).Range("C1").Value = "Tuesday"
    Worksheets(mySheetName).Range("D1").Value = "Wednesday"
    Worksheets(mySheetName).Range("E1").Value = "Thursday"
    Worksheets(mySheetName).Range("F1").Value = "Friday"
    Worksheets(mySheetName).Range("G1").Value = "Saturday"
    Worksheets(mySheetName).Range("A2").Value = sun
    Worksheets(mySheetName).Range("B2").Value = mon
    Worksheets(mySheetName).Range("C2").Value = tues
    Worksheets(mySheetName).Range("D2").Value = wed
    Worksheets(mySheetName).Range("E2").Value = thurs
    Worksheets(mySheetName).Range("F2").Value = fri
    Worksheets(mySheetName).Range("G2").Value = sat

    
End Sub

Sub Weekly_schedule_Milestones()

 startDay = dhFirstDayInWeek()
 endDAy = dhLastDayInWeek()

Call get_weeklyMileStones(startDay, endDAy)


End Sub



