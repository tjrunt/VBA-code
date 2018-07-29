Function WorksheetExists(sheetName As String) As Boolean
    
Dim TempSheetName As String

TempSheetName = UCase(sheetName)


    
For Each sheet In Worksheets
    If TempSheetName = UCase(sheet.Name) Then
        WorksheetExists = True
        Exit Function
    End If
Next sheet
WorksheetExists = False
End Function

Function ganntParty()
    Call makeParametersTemplate
End Function

Sub todo()
    'get min date in col H
    'set first date min formula or script??
    'format date color and angle
    'conditional formatting colors
    'user specifd custom colors????
    'expand date row values
    'apply formula and expand
    'verify data
    'clean up code
    'refactor
    'format cells
    '
    '
    '
    '
    '
    

    'MsgBox ""
    'get font color
    MsgBox Range("c1").Interior.ColorIndex
    MsgBox Range("c8").Interior.ColorIndex
    MsgBox Range("c11").Interior.ColorIndex
    'MsgBox WorksheetExists("myGannt")
    'MsgBox WorksheetExists("XXXmyGannt")

End Sub


''
'' make Templates Code
''

Sub makeParametersTemplate()
    
    
    isWS = WorksheetExists("parameters")
    
    
    If isWS Then
        MsgBox "This Sheet already Exist"
    Else
    
    myparametersSheet = "parameters"
    Sheets.Add.Name = myparametersSheet
        
    Sheets(myparametersSheet).Range("A1").Value = "Gannt options"
    Sheets(myparametersSheet).Range("A2").Value = "SheetName"
    Sheets(myparametersSheet).Range("A3").Value = "Start column"
    Sheets(myparametersSheet).Range("A4").Value = "End column"
    
    Sheets(myparametersSheet).Range("A5").Value = "Task 1 Name"
    Sheets(myparametersSheet).Range("A6").Value = "task 1 start row"
    Sheets(myparametersSheet).Range("A7").Value = "task 1 End row"
    
    Sheets(myparametersSheet).Range("A8").Value = "Task 2 Name"
    Sheets(myparametersSheet).Range("A9").Value = "task 2 start row"
    Sheets(myparametersSheet).Range("A10").Value = "task 2 End row"
    
    Sheets(myparametersSheet).Range("A11").Value = "Task 3 Name"
    Sheets(myparametersSheet).Range("A12").Value = "task 3 start row"
    Sheets(myparametersSheet).Range("A13").Value = "task 3 End row"
    
    Sheets(myparametersSheet).Range("A14").Value = "Task 4 Name"
    Sheets(myparametersSheet).Range("A15").Value = "task 4 start row"
    Sheets(myparametersSheet).Range("A16").Value = "task 4 End row"
    
    Sheets(myparametersSheet).Range("A17").Value = "Task 5 Name"
    Sheets(myparametersSheet).Range("A18").Value = "task 5 start row"
    Sheets(myparametersSheet).Range("A19").Value = "task 5 End row"
    
    Sheets(myparametersSheet).Range("A20").Value = "Task 6 Name"
    Sheets(myparametersSheet).Range("A21").Value = "task 6 start row"
    Sheets(myparametersSheet).Range("A22").Value = "task 6 End row"
    

    Sheets(myparametersSheet).Range("B1").Value = "Enter Parameters Here"


     Sheets(myparametersSheet).Range("G20").Value = "Enter New start Date for Gant Below"
     
     
     Call CustomColorTemplateStuff
     Sheets(myparametersSheet).Range("C1").Value = "Color cells below for Custom Colors in task row"
     Sheets(myparametersSheet).Range("C5").Value = "Task 1 Color"
     Sheets(myparametersSheet).Range("C8").Value = "Task 2 Color"
     Sheets(myparametersSheet).Range("C11").Value = "Task 3 Color"
     Sheets(myparametersSheet).Range("C14").Value = "Task 4 Color"
     Sheets(myparametersSheet).Range("C17").Value = "Task 5 Color"
     Sheets(myparametersSheet).Range("C20").Value = "Task 6 Color"
     
     
    Call formatOptioninoutSheet
    Call addButtonForGannt
    Call addButtonForColorUpdate
    Call addButtonForDateUpdate
    Call addButtonForResetFormula
    
    End If

End Sub

Private Sub makeMyGanntTemplate()
    

    xSheetExist = WorksheetExists("myGannt")
    
    If xSheetExist Then
        Call makeMyGant
        
    Else
    
    myGanntSheetName = "myGannt"
    
       
    Sheets.Add.Name = myGanntSheetName
        
    
    'make value array
    Dim headerRowArray(1 To 19) As String
    headerRowArray(1) = "Ad Date"
    headerRowArray(2) = "Activity ID"
    headerRowArray(3) = "Print Channel"
    headerRowArray(4) = "Stage Names"
    headerRowArray(5) = "Base Page Count"
    headerRowArray(6) = "Total Page Count"
    headerRowArray(7) = "Ad Manager"
    headerRowArray(8) = "Task Start 1"
    headerRowArray(9) = "Task End 1"
    headerRowArray(10) = "Task Start 2"
    headerRowArray(11) = "Task End 2"
    headerRowArray(12) = "Task Start 3"
    headerRowArray(13) = "Task End 3"
    headerRowArray(14) = "Task Start 4"
    headerRowArray(15) = "Task End 4"
    headerRowArray(16) = "Task Start 5"
    headerRowArray(17) = "Task End 5"
    headerRowArray(18) = "Task Start 6"
    headerRowArray(19) = "Task End 6"
    
    'populate forst row with a loop
    Dim i As Integer
    For i = 1 To 19
        Sheets(myGanntSheetName).Cells(1, i).Value = headerRowArray(i)
    Next i
    
        Sheets(myGanntSheetName).Range("T1").Value = Date
        'Call resetMyFormula
        'Call expandFormula
        Call dateFormatting
    
    End If
        

    Call freezeAndColor
    Call makeMyGant
    
End Sub


''
'' button code
''
Private Sub addButtonForGannt()
'
' addButtonForGannt Macro
'

'
    Sheets("parameters").Buttons.Add(621, 83, 120, 30).Select
    Selection.OnAction = "PERSONAL.XLSB!makeMyGanntTemplate"
    Selection.Characters.Text = "Make/update Gannt"
    With Selection.Characters(Start:=1, Length:=20).Font
        .Name = "Arial"
        .FontStyle = "Regular"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    
End Sub

Private Sub addButtonForColorUpdate()
    'todo update function to make sheet active
    Sheets("parameters").Buttons.Add(621, 133, 120, 30).Select
    Selection.OnAction = "PERSONAL.XLSB!removeFormCond"
    Selection.Characters.Text = "Update Colors"
    With Selection.Characters(Start:=1, Length:=20).Font
        .Name = "Arial"
        .FontStyle = "Regular"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With

End Sub

Private Sub addButtonForResetFormula()
    'todo update code for date rest
    Sheets("parameters").Buttons.Add(621, 183, 120, 30).Select
    Selection.OnAction = "PERSONAL.XLSB!resetMyFormula"
    Selection.Characters.Text = "Reset Formula"
    With Selection.Characters(Start:=1, Length:=20).Font
        .Name = "Arial"
        .FontStyle = "Regular"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With

End Sub

Sub updateDate()
    'MsgBox Sheets("parameters").Range("H24").Value
    Sheets("myGannt").Range("T1").Value = Sheets("parameters").Range("H20").Value
        Call dateFormatting



End Sub

Private Sub addButtonForDateUpdate()
    'todo update code for date rest
    Sheets("parameters").Buttons.Add(621, 293, 120, 30).Select
    Selection.OnAction = "PERSONAL.XLSB!updateDate"
    Selection.Characters.Text = "Update Date"
    With Selection.Characters(Start:=1, Length:=20).Font
        .Name = "Arial"
        .FontStyle = "Regular"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With

End Sub
''
''



''
'' Formatting Code
''
Private Sub dateFormatting()
'
' XXXConditionalFormatting Macro
'

'
    Sheets("myGannt").Select
    Range("T1").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 90
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .Name = "Arial"
        .FontStyle = "Regular"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("U1:AE1").Select
    Selection.Insert Shift:=xlToRight
    Range("T1").Select
    Selection.AutoFill Destination:=Range("T1:ZZ1"), Type:=xlFillDefault
    Range("T1:ZZ1").Select
    Columns("T:ZZ").Select
    Columns("T:ZZ").EntireColumn.AutoFit
End Sub

Private Sub freezeAndColor()
'
' freezeAndColor Macro
'

'
    Sheets("myGannt").Select
    Columns("T:T").Select
    ActiveWindow.FreezePanes = True
    Columns("H:I").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.0999786370433668
        .PatternTintAndShade = 0
    End With
    Columns("L:M").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.0999786370433668
        .PatternTintAndShade = 0
    End With
    Columns("P:Q").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.0999786370433668
        .PatternTintAndShade = 0
    End With
    Columns("F:F").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.0999786370433668
        .PatternTintAndShade = 0
    End With
    Columns("D:D").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.0999786370433668
        .PatternTintAndShade = 0
    End With
    Columns("B:B").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.0999786370433668
        .PatternTintAndShade = 0
    End With
    Cells.Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub

Private Sub formatOptioninoutSheet()
'
' formatOptioninoutSheet Macro
'

'
    Columns("A:B").Select
    Range("B1").Activate
    Columns("A:B").EntireColumn.AutoFit
    Range("A1:B1").Select
    Range("B1").Activate
    With Selection.Font
        .Name = "Arial"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    Range("A2:B22").Select
    Range("B2").Activate
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Range("A1:B22").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Columns("A:B").Select
    Range("B1").Activate
    Columns("A:B").EntireColumn.AutoFit
    Selection.ColumnWidth = 16.33
    Selection.ColumnWidth = 27.17
End Sub

''
''
''
Private Sub makeMyGant()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
      
    myGanntSheetName = "myGannt"
    myparametersSheet = "parameters"
     
    'input values start
    sheetName = Sheets(myparametersSheet).Range("B2").Value
    startcolLetter = Sheets(myparametersSheet).Range("B3").Value
    endcolLetter = Sheets(myparametersSheet).Range("B4").Value
    
    task1Name = Sheets(myparametersSheet).Range("B5").Value
    task1startRow = Sheets(myparametersSheet).Range("B6").Value
    task1endRow = Sheets(myparametersSheet).Range("B7").Value
    
    task2Name = Sheets(myparametersSheet).Range("B8").Value
    task2startRow = Sheets(myparametersSheet).Range("B9").Value
    task2endRow = Sheets(myparametersSheet).Range("B10").Value
      
    task3Name = Sheets(myparametersSheet).Range("B11").Value
    task3startRow = Sheets(myparametersSheet).Range("B12").Value
    task3endRow = Sheets(myparametersSheet).Range("B13").Value
    
  
    task4Name = Sheets(myparametersSheet).Range("B14").Value
    task4startRow = Sheets(myparametersSheet).Range("B15").Value
    task4endRow = Sheets(myparametersSheet).Range("B16").Value
    
    task5Name = Sheets(myparametersSheet).Range("B17").Value
    task5startRow = Sheets(myparametersSheet).Range("B18").Value
    task5endRow = Sheets(myparametersSheet).Range("B19").Value
    
    task6Name = Sheets(myparametersSheet).Range("B20").Value
    task6startRow = Sheets(myparametersSheet).Range("B21").Value
    task6endRow = Sheets(myparametersSheet).Range("B22").Value
           

    'input values End

    Dim StartColumnNumber As Long
    Dim EndColumnNumber As Long
    ''todo test & fix error
    'Convert To Column Number
    StartColumnNumber = Range(startcolLetter & 1).Column
    EndColumnNumber = Range(endcolLetter & 1).Column

    firstEmptyRowinColA = Sheets(myGanntSheetName).Cells(Rows.count, 1).End(xlUp).Row + 1
    
    'static values from parameters sheet
    printChanelVal = sheetName
    stagNamesVal = task1Name & " | " & task2Name & " | " & task3Name & " | " & task4Name & " | " & task5Name & " | " & task6Name

'loop get and set gant values
destRow = firstEmptyRowinColA

For colStart = StartColumnNumber To EndColumnNumber
    adDateVal = Sheets(sheetName).Cells(3, colStart).Value
    x = Left(adDateVal, 1)
    If (IsNumeric(x) = True) Then
        
    
        'get values
        adDateVal = Sheets(sheetName).Cells(3, colStart).Value
        ActIDVal = Sheets(sheetName).Cells(6, colStart).Value

        basePGCountVal = Sheets(sheetName).Cells(8, colStart).Value
        adManagerVal = Sheets(sheetName).Cells(10, colStart).Value
        
        ''todo needs if not empty
        If (task1startRow <> "") Then
            t1StartVal = Sheets(sheetName).Cells(task1startRow, colStart).Value
            t1EndVal = Sheets(sheetName).Cells(task1endRow, colStart).Value
        End If
        If (task2startRow <> "") Then
            t2StartVal = Sheets(sheetName).Cells(task2startRow, colStart).Value
            t2EndVal = Sheets(sheetName).Cells(task2endRow, colStart).Value
        End If
        If (task3startRow <> "") Then
            t3StartVal = Sheets(sheetName).Cells(task3startRow, colStart).Value
            t3EndVal = Sheets(sheetName).Cells(task3endRow, colStart).Value
        End If
        If (task4startRow <> "") Then
            t4StartVal = Sheets(sheetName).Cells(task4startRow).Value
            t4EndVal = Sheets(sheetName).Cells(task4endRow, colStart).Value
        End If
        If (task5startRow <> "") Then
            t5StartVal = Sheets(sheetName).Cells(task5startRow, colStart).Value
            t5EndVal = Sheets(sheetName).Cells(task5endRow, colStart).Value
        End If
        If (task6startRow <> "") Then
            t6StartVal = Sheets(sheetName).Cells(task6startRow, colStart).Value
            t6EndVal = Sheets(sheetName).Cells(task6endRow, colStart).Value
        End If
        
        
        
           

        'populate Gannt cell
        Sheets(myGanntSheetName).Cells(destRow, 1).Value = adDateVal
        Sheets(myGanntSheetName).Cells(destRow, 2).Value = ActIDVal
        Sheets(myGanntSheetName).Cells(destRow, 3).Value = printChanelVal
        Sheets(myGanntSheetName).Cells(destRow, 4).Value = stagNamesVal
        Sheets(myGanntSheetName).Cells(destRow, 5).Value = basePGCountVal
        '
        '
        Sheets(myGanntSheetName).Cells(destRow, 7).Value = adManagerVal
        Sheets(myGanntSheetName).Cells(destRow, 8).Value = t1StartVal
        Sheets(myGanntSheetName).Cells(destRow, 9).Value = t1EndVal
        Sheets(myGanntSheetName).Cells(destRow, 10).Value = t2StartVal
        Sheets(myGanntSheetName).Cells(destRow, 11).Value = t2EndVal
        Sheets(myGanntSheetName).Cells(destRow, 12).Value = t3StartVal
        Sheets(myGanntSheetName).Cells(destRow, 13).Value = t3EndVal
        Sheets(myGanntSheetName).Cells(destRow, 14).Value = t4StartVal
        Sheets(myGanntSheetName).Cells(destRow, 15).Value = t4EndVal
        Sheets(myGanntSheetName).Cells(destRow, 16).Value = t5StartVal
        Sheets(myGanntSheetName).Cells(destRow, 17).Value = t5EndVal
        Sheets(myGanntSheetName).Cells(destRow, 18).Value = t6StartVal
        Sheets(myGanntSheetName).Cells(destRow, 19).Value = t6EndVal

        
        destRow = destRow + 1
    End If
    
Next colStart
  
    
    ''myformula
    '=IF(AND(T$1>=$H2,T$1<=$I2),"1X", IF(AND(T$1>=$J2,T$1<=$K2),"2X", IF(AND(T$1>=$L2,T$1<=$M2),"3X",IF(AND(T$1>=$N2,T$1<=$O2),"4X", IF(AND(T$1>=$P2,T$1<=$Q2),"5X", IF(AND(T$1>=$R2,T$1<=$S2),"6X",""))))))
    
    
    Call setMinDate
    
    Call resetMyFormula
    
    
    Call removeFormCond
    
    Call AddFormCond
    
    MsgBox "Done"

End Sub


''
''
''

Private Sub expandFormula()
 

    Sheets("myGannt").Select
    Application.CutCopyMode = False
    Selection.AutoFill Destination:=Range("T2:T675"), Type:=xlFillDefault
    Range("T2:T675").Select
    Selection.AutoFill Destination:=Range("T2:AAG675"), Type:=xlFillDefault
    Range("T2:AAG675").Select
   
    

 End Sub
 
Private Sub resetMyFormula()
'
    Sheets("myGannt").Select
    ActiveCell.FormulaR1C1 = ""
    Range("T2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(R1C>=RC8,R1C<=RC9),""1X"", IF(AND(R1C>=RC10,R1C<=RC11),""2X"", IF(AND(R1C>=RC12,R1C<=RC13),""3X"",IF(AND(R1C>=RC14,R1C<=RC15),""4X"", IF(AND(R1C>=RC16,R1C<=RC17),""5X"", IF(AND(R1C>=RC18,R1C<=RC19),""6X"",""""))))))"
    
    Call expandFormula
    
End Sub

Private Sub setMinDate()

     myGanntSheetName = "myGannt"
     
     If IsDate(Sheets(myGanntSheetName).Range("H2").Value) Then
        Sheets(myGanntSheetName).Range("T1").Value = Sheets(myGanntSheetName).Range("H2").Value
        Call dateFormatting
        
    End If
    
End Sub



''Crap
'http://www.bluepecantraining.com/portfolio/excel-vba-macro-to-apply-conditional-formatting-based-on-value/


Sub removeFormCond()
        Dim rg As Range

     Sheets("myGannt").Select
    Set rg = Range("T2:AAG675")

    rg.FormatConditions.Delete
    'todo Add custom color Functions
    
    
End Sub

Sub AddFormCond()
    'http://www.bluepecantraining.com/portfolio/excel-vba-macro-to-apply-conditional-formatting-based-on-value/
    Dim rg As Range
    
    'if not -4142
    
	colorX1 = Sheets("parameters").Select.Range("c5").Interior.ColorIndex
	colorX2 = Sheets("parameters").Select.Range("c8").Interior.ColorIndex
	colorX3 = Sheets("parameters").Select.Range("c11").Interior.ColorIndex
	colorX4 = Sheets("parameters").Select.Range("c14").Interior.ColorIndex
	colorX5 = Sheets("parameters").Select.Range("c17").Interior.ColorIndex
	colorX6 = Sheets("parameters").Select.Range("c20").Interior.ColorIndex
    

     Sheets("myGannt").Select
    Set rg = Range("T2:AAG675")
    
   
    rg.FormatConditions.Delete
    With rg.FormatConditions.Add( _
        Type:=xlCellValue, _
        Operator:=xlEqual, _
        Formula1:="1X")
    
        .Interior.Color = RGB(198, 66, 66)
        .Font.Color = RGB(198, 66, 66)
    
    End With


    With rg.FormatConditions.Add( _
        Type:=xlCellValue, _
        Operator:=xlEqual, _
        Formula1:="2X")
    
        .Interior.Color = RGB(150, 2, 206)
        .Font.Color = RGB(150, 2, 206)
    
    End With
    
        With rg.FormatConditions.Add( _
        Type:=xlCellValue, _
        Operator:=xlEqual, _
        Formula1:="2X")
    
        .Interior.Color = RGB(198, 0, 150)
        .Font.Color = RGB(198, 0, 150)
    
    End With
    
        With rg.FormatConditions.Add( _
        Type:=xlCellValue, _
        Operator:=xlEqual, _
        Formula1:="3X")
    
        .Interior.Color = RGB(44, 6, 180)
        .Font.Color = RGB(44, 6, 180)
    
    End With
    
    
        With rg.FormatConditions.Add( _
        Type:=xlCellValue, _
        Operator:=xlEqual, _
        Formula1:="4X")
    
        .Interior.Color = RGB(130, 150, 4)
        .Font.Color = RGB(130, 150, 4)
    
    End With
    
    
        With rg.FormatConditions.Add( _
        Type:=xlCellValue, _
        Operator:=xlEqual, _
        Formula1:="5X")
    
        .Interior.Color = RGB(55, 55, 100)
        .Font.Color = RGB(55, 55, 100)
    
    End With

        With rg.FormatConditions.Add( _
        Type:=xlCellValue, _
        Operator:=xlEqual, _
        Formula1:="6X")
    
        .Interior.Color = RGB(33, 133, 133)
        .Font.Color = RGB(33, 133, 133)
    
    End With



End Sub






'todo add optional header info from rows
'todo fix first date
'todo add color update button
'todo add update dates button + formla
'todo fix blanks

'todo get and set condition formatting colors
Sub getCellColor()

	'if not -4142
	color1 = Sheets("parameters").Select.Range("c5").Interior.ColorIndex
	color2 = Sheets("parameters").Select.Range("c8").Interior.ColorIndex
	color3 = Sheets("parameters").Select.Range("c11").Interior.ColorIndex
	color4 = Sheets("parameters").Select.Range("c14").Interior.ColorIndex
	color5 = Sheets("parameters").Select.Range("c17").Interior.ColorIndex
	color6 = Sheets("parameters").Select.Range("c20").Interior.ColorIndex
End Sub







