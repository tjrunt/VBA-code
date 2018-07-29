Attribute VB_Name = "hlDAtesCode"


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

Private Sub kill_color()
    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets

    For Each Cell In ws.UsedRange.Cells
        If IsDate(Cell.Value) Then
        
        If Cell.Interior.colorIndex = 43 Then
            Cell.Interior.colorIndex = 0
           ElseIf Cell.Interior.colorIndex = 44 Then
            Cell.Interior.colorIndex = 0
        ElseIf Cell.Interior.colorIndex = 42 Then
            Cell.Interior.colorIndex = 0
        ElseIf Cell.Interior.colorIndex = 46 Then
            Cell.Interior.colorIndex = 0
            
            End If
   
   End If
     Next Cell
Next ws
    
End Sub



Private Sub dateRange_kill_color()
    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets

    For Each Cell In ws.UsedRange.Cells
        myCell = Cell.Value
                 If IsError(Cell.Value) Then
            If Cell.Value = CVErr(2023) Then
            'do nothing, e.g.:
            x = 0
        End If
        
        ElseIf (Len(myCell) <> 0) And (Len(myCell) <= 11) And (InStr(1, myCell, "/") > 0) And (InStr(1, myCell, "-") > 0) Then
        
        If Cell.Interior.colorIndex = 43 Then
            Cell.Interior.colorIndex = 0
           ElseIf Cell.Interior.colorIndex = 44 Then
            Cell.Interior.colorIndex = 0
        ElseIf Cell.Interior.colorIndex = 42 Then
            Cell.Interior.colorIndex = 0
        ElseIf Cell.Interior.colorIndex = 46 Then
            Cell.Interior.colorIndex = 0
            
            End If
   
   End If
     Next Cell
Next ws
    
End Sub

Private Sub highlight_CurrentWeek_MileStones(startDate, endDate, colorValue)
    On Error Resume Next
     Dim ws As Worksheet
     
startDay = startDate
endDAy = endDate

For Each ws In ActiveWorkbook.Worksheets
    ws.Visible = xlSheetVisible
    For Each Cell In ws.UsedRange.Cells
         'do some stuff
         
         If IsError(Cell.Value) Then
            If Cell.Value = CVErr(2023) Then
            'do nothing, e.g.:
            x = 0
        End If
         
         ElseIf (Cell.Value >= startDay) And (Cell.Value <= endDAy) And (Cell.Font.colorIndex = 1) Then
             Cell.Interior.colorIndex = colorValue

        End If
   
     Next Cell
Next ws
   

End Sub

Private Sub highlight_MultiWeek_MileStones(startDate, endDate, colorValue)
     Dim ws As Worksheet
    startDay = startDate
    endDAy = endDate
On Error GoTo eh
For Each ws In ActiveWorkbook.Worksheets
    ws.Visible = xlSheetVisible
    For Each Cell In ws.UsedRange.Cells
        myCell = Cell.Value
         If IsError(Cell.Value) Then
            If Cell.Value = CVErr(2023) Then
            'do nothing, e.g.:
            x = 0
        End If
        
        ElseIf (Len(myCell) <> 0) And (Len(myCell) <= 11) And (InStr(1, myCell, "/") > 0) And (InStr(1, myCell, "-") > 0) Then
        mycellSplit = Split(myCell, "-")
        mydate = mycellSplit(0)
            
            If (CDate(mydate) >= startDay) And (CDate(mydate) <= endDAy) And (Cell.Font.colorIndex = 1) Then
                Cell.Interior.colorIndex = colorValue
            End If
        
        End If
     Next Cell
Next ws
eh:
    MsgBox "The following error occurred: " & Err.Description & " cell value is: " & myCell


End Sub


Sub resetColors()

Call kill_color
Call dateRange_kill_color

MsgBox " Done"


End Sub


Sub AA_HLthisWeekPlusThree()

startDay = dhFirstDayInWeek()
endDAy = dhLastDayInWeek()

Call highlight_CurrentWeek_MileStones(startDay, endDAy, 43)
Call highlight_CurrentWeek_MileStones(startDay + 7, endDAy + 7, 44)
Call highlight_CurrentWeek_MileStones(startDay + 14, endDAy + 14, 42)
Call highlight_CurrentWeek_MileStones(startDay + 21, endDAy + 21, 46)

MsgBox "Done"


End Sub

Sub AA_Ranges_HLthisWeekPlusThree()

startDay = dhFirstDayInWeek()
endDAy = dhLastDayInWeek()

Call highlight_MultiWeek_MileStones(startDay, endDAy, 43)
Call highlight_MultiWeek_MileStones(startDay + 7, endDAy + 7, 44)
Call highlight_MultiWeek_MileStones(startDay + 14, endDAy + 14, 42)
Call highlight_MultiWeek_MileStones(startDay + 21, endDAy + 21, 46)

MsgBox "Done"


End Sub




Sub BB_HLnextWeekPlusThree()

startDay = dhFirstDayInWeek()
endDAy = dhLastDayInWeek()

Call highlight_MultiWeek_MileStones(startDay + 7, endDAy + 7, 43)
Call highlight_MultiWeek_MileStones(startDay + 14, endDAy + 14, 44)
Call highlight_MultiWeek_MileStones(startDay + 21, endDAy + 21, 42)
Call highlight_MultiWeek_MileStones(startDay + 28, endDAy + 28, 46)



MsgBox "Done"

End Sub




Sub BB_Ranges_HLnextWeekPlusThree()

startDay = dhFirstDayInWeek()
endDAy = dhLastDayInWeek()

Call highlight_MultiWeek_MileStones(startDay + 7, endDAy + 7, 43)
Call highlight_MultiWeek_MileStones(startDay + 14, endDAy + 14, 44)
Call highlight_MultiWeek_MileStones(startDay + 21, endDAy + 21, 42)
Call highlight_MultiWeek_MileStones(startDay + 28, endDAy + 28, 46)



MsgBox "Done"

End Sub




Private Sub xxxTestMultiWeek()

'36 6  14
MsgBox Range("AE14").Font.colorIndex
'startDay = dhFirstDayInWeek()
'endDAy = dhLastDayInWeek()

'update kill color for multiweek
'Call kill_color
'Call dateRange_kill_color
'Call highlight_MultiWeek_MileStones(startDay, endDAy, 43)
'Call highlight_MultiWeek_MileStones(startDay + 7, endDAy + 7, 44)
'Call highlight_MultiWeek_MileStones(startDay + 14, endDAy + 14, 42)
'Call highlight_MultiWeek_MileStones(startDay + 21, endDAy + 21, 46)


'MsgBox "Done"


End Sub









