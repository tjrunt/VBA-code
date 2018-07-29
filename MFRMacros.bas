Attribute VB_Name = "MFRs"
Private Sub incrementAlphabet(PageName)

Dim WrdArray() As String
sourceCellNum = 2

myCell = "I" & sourceCellNum
alphas = Array("A", "B", "C", "D", "E", "F", "G", "H", "J", "K", "L", "M", "N", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA")

Dim newFolio As String
Dim myNextLetter As String

Do While Len(Worksheets(PageName).Range(myCell).Value) > 0
    
    x = Worksheets(PageName).Range(myCell).Value
    WrdArray() = Split(x, "_")
    myLetter = Right(WrdArray(0), 1)
    If IsNumeric(myLetter) Then
    
        newFolio = WrdArray(0) & "A" & "_" & WrdArray(1) & "_" & WrdArray(2)
        Worksheets(PageName).Range(myCell).Value = newFolio
    
    Else


    For i = 0 To 24
        If myLetter = alphas(i) Then
        
                index = i + 1
                
                'rlpStr = StrReverse(Replace(StrReverse(str), StrReverse("Help"),  StrReverse("Job"), , 1))
                'newLetter = Replace(WrdArray(0), myLetter, alphas(index))
                newLetter = StrReverse(Replace(StrReverse(WrdArray(0)), StrReverse(myLetter), StrReverse(alphas(index)), , 1))
                newFolio = newLetter & "_" & WrdArray(1) & "_" & WrdArray(2)
                
                Worksheets(PageName).Range(myCell).Value = newFolio
            End If
        Next i
  End If
  
  sourceCellNum = sourceCellNum + 1
  myCell = "I" & sourceCellNum
  Loop

End Sub

Sub alaska_frm_BaseMVR()
    Sheets("PivotTable").Select
    myaddate = InputBox("Enter four digit Folio ad date e.x. 1012 or 0213")

    myactID = InputBox("Enter activity ID")
    
    
    myMc = 608
    myMn = "Alaska-1"
    myMa = "MA23"
    sa = "SA"
    Call addWS
    
     'todo  header row formatting
     
     Call offshoreTemplateItems
   Call getValues_offshore_from_base
   pgn = "sheet1"
   Call incrementAlphabet(pgn)
   'todo call & populate alaska values
   'save and other crap
   
   Call v2getValues_offshore("I", 2, "J", 2, "Sheet1", myaddate, myactID, myMc, myMn, myMa, sa)
   Call saveSheetToNewBook

   
End Sub

Sub hawaii_frm_BaseMVR()
    Sheets("PivotTable").Select
    myaddate = InputBox("Enter four digit Folio ad date e.x. 1012 or 0213")

    myactID = InputBox("Enter activity ID")
    
    
    haMc = 745
    haMn = "Hawaii-1"
    haMa = "MA22"
    Sh = "SH"
    Call addWS
    
     'todo  header row formatting
     
     Call offshoreTemplateItems
   Call getValues_offshore_from_base
   pgn = "sheet1"
   Call incrementAlphabet(pgn)
   Call incrementAlphabet(pgn)
   'todo call & populate alaska values
   'save and other crap
   
   Call v2getValues_offshore("I", 2, "J", 2, "Sheet1", myaddate, myactID, haMc, haMn, haMa, Sh)
   Call saveSheetToNewBook

   
End Sub

Private Sub offshoreTemplateItems()
'todo ad color for cells


Worksheets("Sheet1").Range("A1").Value = "ActivityId"
Worksheets("Sheet1").Range("B1").Value = "VersionId"
Worksheets("Sheet1").Range("C1").Value = "MarketCode"
Worksheets("Sheet1").Range("D1").Value = "MarketName"
Worksheets("Sheet1").Range("E1").Value = "MerchArea"
Worksheets("Sheet1").Range("F1").Value = "PageID"
Worksheets("Sheet1").Range("G1").Value = "PageSequenceID"
Worksheets("Sheet1").Range("H1").Value = "WorkingPageID"
Worksheets("Sheet1").Range("I1").Value = "PageName"
Worksheets("Sheet1").Range("J1").Value = "PageFolio"

''
''

Worksheets("Sheet1").Range("A1").Interior.Color = RGB(0, 0, 255)
Worksheets("Sheet1").Range("B1").Interior.Color = RGB(0, 0, 255)
Worksheets("Sheet1").Range("C1").Interior.Color = RGB(0, 0, 255)
Worksheets("Sheet1").Range("D1").Interior.Color = RGB(0, 0, 255)
Worksheets("Sheet1").Range("E1").Interior.Color = RGB(0, 0, 255)
Worksheets("Sheet1").Range("F1").Interior.Color = RGB(0, 0, 255)
Worksheets("Sheet1").Range("G1").Interior.Color = RGB(0, 0, 255)
Worksheets("Sheet1").Range("H1").Interior.Color = RGB(0, 0, 255)
Worksheets("Sheet1").Range("I1").Interior.Color = RGB(0, 0, 255)
Worksheets("Sheet1").Range("J1").Interior.Color = RGB(0, 0, 255)


Worksheets("Sheet1").Range("A1").Font.Color = RGB(255, 255, 255)
Worksheets("Sheet1").Range("B1").Font.Color = RGB(255, 255, 255)
Worksheets("Sheet1").Range("C1").Font.Color = RGB(255, 255, 255)
Worksheets("Sheet1").Range("D1").Font.Color = RGB(255, 255, 255)
Worksheets("Sheet1").Range("E1").Font.Color = RGB(255, 255, 255)
Worksheets("Sheet1").Range("F1").Font.Color = RGB(255, 255, 255)
Worksheets("Sheet1").Range("G1").Font.Color = RGB(255, 255, 255)
Worksheets("Sheet1").Range("H1").Font.Color = RGB(255, 255, 255)
Worksheets("Sheet1").Range("I1").Font.Color = RGB(255, 255, 255)
Worksheets("Sheet1").Range("J1").Font.Color = RGB(255, 255, 255)


End Sub

'todo spin off hawaii
Private Sub v2getValues_offshore(sourceColLetter, sourceCellNum, destColLetter, destCellNum, PageName, addate, actID, mc, mn, ma, adFormat)

Dim myCell As String
Dim pgName As String
Dim WrdArray() As String
Dim text_string As String
addate = addate
actID = actID


pgName = PageName
sourceCellNum = sourceCellNum
destCellNum = destCellNum
currentCell = sourceColLetter & sourceCellNum
destCell = destColLetter & destCellNum
actIDcell = "A" & destCellNum
MarketCodeCell = "C" & destCellNum
MarketNameCell = "D" & destCellNum
MerchAreaCell = "E" & destCellNum


Do While Len(Worksheets(pgName).Range(currentCell).Value) > 0
    
    x = Worksheets(pgName).Range(currentCell).Value
    WrdArray() = Split(x, "_")
    
    folioBegining = Left(WrdArray(0), 2)
    
    If Left(folioBegining, 1) = "0" Then
        folioBegining = Replace(folioBegining, "0", "")
    End If
    
    pgNameVal = folioBegining & " " & UCase(adFormat) & addate & WrdArray(1) & WrdArray(2) & "_" & WrdArray(0)
    
    Worksheets(pgName).Range(destCell).Value = pgNameVal
    Worksheets(pgName).Range(actIDcell).Value = actID
    Worksheets(pgName).Range(MarketCodeCell).Value = mc
    Worksheets(pgName).Range(MarketNameCell).Value = mn
    Worksheets(pgName).Range(MerchAreaCell).Value = ma
    
    sourceCellNum = sourceCellNum + 1
    destCellNum = destCellNum + 1
    currentCell = sourceColLetter & sourceCellNum
    destCell = destColLetter & destCellNum
    actIDcell = "A" & destCellNum
    MarketCodeCell = "C" & destCellNum
    MarketNameCell = "D" & destCellNum
    MerchAreaCell = "E" & destCellNum
  Loop


End Sub

Private Sub getValues_offshore_from_base()

Dim myCell As String
Dim pgName As String
pgName = "PivotTable"


colA = "A"
colB = "B"
colD = "D"
startCellNum = 3

colA_destination = "G"
colB_destination = "H"
colD_destination = "I"
cell_dest = 2
cell_dest_plust_one = startCellNum + 1
A_cell_onedown = "A" & cell_dest_plust_one

gt = "Grand Total"

PageNamecurrentcell = colD & startCellNum
CheckCell = ""




 Do While Len(Worksheets(pgName).Range(PageNamecurrentcell).Value) > 0
    ' If Len(Worksheets(pgName).Range(A_cell_onedown).Value) > 0 And Worksheets(pgName).Range(A_cell_onedown).Value <> gt Then
    If Len(Worksheets(pgName).Range(A_cell_onedown).Value) > 0 Then
        ' todo populate cells with highes art alts
    

        Worksheets("Sheet1").Range(colD_destination & cell_dest).Value = Worksheets(pgName).Range(PageNamecurrentcell).Value
        cell_dest = cell_dest + 1
        End If
        
        
        startCellNum = startCellNum + 1
        cell_dest_plust_one = startCellNum + 1
        PageNamecurrentcell = colD & startCellNum
        A_cell_onedown = "A" & cell_dest_plust_one
        
  Loop




'new loop populate with col A & B
startCellNum = 3
cell_dest = 2
 Do While Len(Worksheets(pgName).Range(colD & startCellNum).Value) > 0
    ' todo populate cells with  squid and WPID
    If Len(Worksheets(pgName).Range(colA & startCellNum).Value) > 0 Then
        'populatte offshore mvr

        
        Worksheets("Sheet1").Range(colA_destination & cell_dest).Value = Worksheets(pgName).Range(colA & startCellNum).Value
        Worksheets("Sheet1").Range(colB_destination & cell_dest).Value = Worksheets(pgName).Range(colB & startCellNum).Value
        
        
         cell_dest = cell_dest + 1
    End If
    
    startCellNum = startCellNum + 1
   
    Loop


End Sub

Sub Offshore_MFR()
''''
mySheet = ActiveSheet.Name
Sheets(mySheet).Name = "MVR"
mySheet = ActiveSheet.Name
Call addWS
Call addate
Call templateItems


' populate MFR Market ID (in OM) col E
Call getValues("C", "2", "E", "7", mySheet)
'populate MFR  File (PDF) Name (for Proof) cod D
Call getValues("I", 2, "D", 7, mySheet)
'populate MFR  Page Name(Master folio/Prose) Col C
Call getValues("I", 2, "C", 7, mySheet)
'MFR getpageID  col B
Call getValues("F", 2, "B", 7, mySheet)
'date
Worksheets("Sheet1").Range("H1").Value = Date
' act id
Worksheets("Sheet1").Range("C3").Value = Worksheets(mySheet).Range("A2").Value
'MFR storeformat s or k
Call offshore_sears_or_kmart(mySheet)
'MRF AD Date J2

'MFR Squid  col A
Call getValues("G", 2, "A", 7, mySheet)


'''''
'MFR save as single sheet and name it
Call saveSheetToNewBook



End Sub


Sub Base_MFR()
''''
Sheets("PivotTable").Select
Call addWS
Call addate
Call templateItems

' populate MFR Market ID (in OM) col E
Call getValues("B", "2", "E", "7", "PageName-Market")
'populate MFR  File (PDF) Name (for Proof) cod D
Call getValues("A", 2, "D", 7, "PageName-Market")
'populate MFR  Page Name(Master folio/Prose) Col C
Call getValues("A", 2, "C", 7, "PageName-Market")
'MFR getpageID  col B
Call getpageID("A", 2, "B", 7, "PageName-Market")
'date
Worksheets("Sheet1").Range("H1").Value = Date
' act id
Worksheets("Sheet1").Range("C3").Value = Worksheets(3).Range("A2").Value
'MFR storeformat s or k
Call sears_or_kmart
'MRF AD Date J2

'MFR Squid  col A
Call getSquid("A", 2, "C", 2)


'''''
'MFR save as single sheet and name it
Call saveSheetToNewBook



End Sub

Private Sub offshore_sears_or_kmart(PageName)
'
'
Dim myVal, mySplit() As String

myVal = Worksheets(PageName).Range("J2").Value
mySplit() = Split(myVal, " ")


If Left(mySplit(1), 1) = "S" Then
            
Worksheets("Sheet1").Range("C2").Value = "Sears"
ElseIf Left(mySplit(1), 1) = "K" Then
            
Worksheets("Sheet1").Range("C2").Value = "Kmart"

Else
Worksheets("Sheet1").Range("C2").Value = ""
End If


End Sub
 Private Sub addWS()
Dim ws As Worksheet

     Set ws = Sheets.Add(after:=Sheets(Worksheets.count))
     ws.Name = "Sheet1"
    

 End Sub
 
 Private Sub addWSname(sheetName)
Dim ws As Worksheet

     Set ws = Sheets.Add(after:=Sheets(Worksheets.count))
     ws.Name = sheetName
    

 End Sub
 
 
 
 
Private Sub getSquid(squidColLetter, squidCellNum, pgIDColLetter, pgIDCellNum)
    Dim squid() As String
    Dim pgid() As String
    Dim count As Integer
    
    currentsquidCell = squidColLetter & CStr(squidCellNum)
    currentPgIdCell = pgIDColLetter & CStr(pgIDCellNum)
    count = 0


 Do While Worksheets(1).Range(currentsquidCell).Value <> "Grand Total"
    currentsquidCell = squidColLetter & CStr(squidCellNum)
    currentPgIdCell = pgIDColLetter & CStr(pgIDCellNum)

        If Len(Worksheets(1).Range(currentsquidCell).Value) > 0 And Worksheets(1).Range(currentsquidCell).Value <> "Grand Total" Then

            
            ReDim Preserve squid(0 To count + 1) As String
            ReDim Preserve pgid(0 To count + 1) As String
            
            squid(count) = Worksheets(1).Range(currentsquidCell).Value
            pgid(count) = Worksheets(1).Range(currentPgIdCell).Value
            
            count = count + 1
    
    
    End If

    squidCellNum = squidCellNum + 1
    pgIDCellNum = pgIDCellNum + 1
    Loop
    

    Dim index As Integer
    Dim cellNum As Integer
    
    cellNum = 7
    getcurrentSquidCell = "A" & CStr(cellNum)
    getcurrentPgIdCell = "B" & CStr(cellNum)
    
   
    lastRowInB = Cells(Rows.count, 2).End(xlUp).Row
    i = 7
   Do While i <= lastRowInB
        i = i + 1

    index = 0
    cellNum = cellNum + 1
    For Each Item In pgid
        index = index + 1
        Dim myCellVal As String
        Dim myItem As String
        myCellVal = Worksheets("Sheet1").Range(getcurrentPgIdCell).Value
        myItem = Item
        If myItem = myCellVal Then
        
            Worksheets("Sheet1").Range(getcurrentSquidCell).Value = squid(index - 1)

        End If
    Next
        getcurrentSquidCell = "A" & CStr(cellNum)
    getcurrentPgIdCell = "B" & CStr(cellNum)
 Loop
    
End Sub

Private Sub templateItems()

Worksheets("Sheet1").Range("A1").Value = "Market Version Report (MVR)"
Worksheets("Sheet1").Range("A2").Value = "Store Format"
Worksheets("Sheet1").Range("A3").Value = "Activity Number"
Worksheets("Sheet1").Range("A4").Value = "Ad Date"
Worksheets("Sheet1").Range("A6").Value = "Page Sequence(Source-OM)"

Worksheets("Sheet1").Range("B6").Value = "Page ID(Source-OM)"

Worksheets("Sheet1").Range("C6").Value = "Page Name(Master folio/Prose)"
Worksheets("Sheet1").Range("D6").Value = "File (PDF) Name (for Proof)"
Worksheets("Sheet1").Range("E6").Value = "Market ID (in OM)"
Worksheets("Sheet1").Range("F6").Value = "Page folio (On Proof)"
Worksheets("Sheet1").Range("G6").Value = "Market description (in OM)"
Worksheets("Sheet1").Range("H6").Value = "City/State/format/store # (not on Proof_info from Planner)"

Worksheets("Sheet1").Range("G1").Value = "DATE Created/updated:"

Worksheets("Sheet1").Range("D3").Value = "Channel Type"
Worksheets("Sheet1").Range("E3").Value = "Circular"

End Sub




Private Sub addate()
Dim val As String
val = InputBox("Enter the ad date")
Worksheets("Sheet1").Range("C4").Value = val
End Sub


Private Sub sears_or_kmart()
'
'
Dim myVal, mySplit() As String

myVal = Worksheets(3).Range("J2").Value
mySplit() = Split(myVal, " ")


If Left(mySplit(1), 1) = "S" Then
            
Worksheets("Sheet1").Range("C2").Value = "Sears"
ElseIf Left(mySplit(1), 1) = "K" Then
            
Worksheets("Sheet1").Range("C2").Value = "Kmart"

Else
Worksheets("Sheet1").Range("C2").Value = ""
End If


End Sub



Private Sub getpageID(sourceColLetter, sourceCellNum, destColLetter, destCellNum, PageName)
'
'
Dim myVal, mySplit() As String

sourceColLetter = sourceColLetter
sourceCellNum = sourceCellNum
destColLetter = destColLetter
sdestCellNum = sdestCellNum


currentCell = sourceColLetter & sourceCellNum
destCell = destColLetter & destCellNum


 Do While Len(Worksheets(PageName).Range(currentCell).Value) > 0
    myVal = Worksheets(PageName).Range(currentCell).Value
    mySplit() = Split(myVal, "_")
    Worksheets("Sheet1").Range(destCell).Value = mySplit(2)
    sourceCellNum = sourceCellNum + 1
    destCellNum = destCellNum + 1
    currentCell = sourceColLetter & sourceCellNum
    destCell = destColLetter & destCellNum
      
      Loop







End Sub

Private Sub getValues(sourceColLetter, sourceCellNum, destColLetter, destCellNum, PageName)
'

Dim myCell As String
Dim pgName As String
pgName = PageName
sourceCellNum = sourceCellNum
destCellNum = destCellNum
currentCell = sourceColLetter & sourceCellNum
destCell = destColLetter & destCellNum




 Do While Len(Worksheets(pgName).Range(currentCell).Value) > 0
    
    
    Worksheets("Sheet1").Range(destCell).Value = Worksheets(pgName).Range(currentCell).Value
    sourceCellNum = sourceCellNum + 1
    destCellNum = destCellNum + 1
    currentCell = sourceColLetter & sourceCellNum
    destCell = destColLetter & destCellNum
  Loop


End Sub

Private Sub saveSheetToNewBook()
'
' saveSheetToNewBook Macro
'

'
    Sheets("Sheet1").Select
    Sheets("Sheet1").Move

    'ActiveWorkbook.SaveAs
        'xlOpenXMLWorkbook , CreateBackup:=False
 
        Application.Dialogs(xlDialogSaveAs).Show

End Sub



Private Sub getValues_offshore(sourceColLetter, sourceCellNum, destColLetter, destCellNum, PageName)

Dim myCell As String
Dim pgName As String
Dim pgn As String
Dim WrdArray() As String
Dim text_string As String

pgName = PageName
sourceCellNum = sourceCellNum
destCellNum = destCellNum
currentCell = sourceColLetter & sourceCellNum
destCell = destColLetter & destCellNum
actIDcell = "A" & destCellNum
MarketCodeCell = "C" & destCellNum
MarketNameCell = "D" & destCellNum
MerchAreaCell = "E" & destCellNum

addate = InputBox("Enter four digit Folio ad date e.x. 1012 or 0213")
adFormat = InputBox("Enter ad folio format e.x. sears Alaska = SA  Hawaii = SH ect...")
actID = InputBox("Enter activity ID")
actID = actID
adFormat = UCase(adFormat)

pgn = "MasterData"
Call incrementAlphabet(pgn)

If Left(adFormat, 1) = "S" And Right(adFormat, 1) = "A" Then
    mc = 608
    mn = "Alaska-1"
    ma = "MA23"

ElseIf Left(adFormat, 1) = "S" And Right(adFormat, 1) = "H" Then
    mc = 745
    mn = "Hawaii-1"
    ma = "MA22"

End If


Do While Len(Worksheets(pgName).Range(currentCell).Value) > 0
    
    x = Worksheets(pgName).Range(currentCell).Value
    WrdArray() = Split(x, "_")
    
    folioBegining = Left(WrdArray(0), 2)
    
    If Left(folioBegining, 1) = "0" Then
        folioBegining = Replace(folioBegining, "0", "")
    End If
    
    pgNameVal = folioBegining & " " & UCase(adFormat) & addate & WrdArray(1) & WrdArray(2) & "_" & WrdArray(0)
    
    Worksheets(pgName).Range(destCell).Value = pgNameVal
    Worksheets(pgName).Range(actIDcell).Value = actID
    Worksheets(pgName).Range(MarketCodeCell).Value = mc
    Worksheets(pgName).Range(MarketNameCell).Value = mn
    Worksheets(pgName).Range(MerchAreaCell).Value = ma
    
    sourceCellNum = sourceCellNum + 1
    destCellNum = destCellNum + 1
    currentCell = sourceColLetter & sourceCellNum
    destCell = destColLetter & destCellNum
    actIDcell = "A" & destCellNum
    MarketCodeCell = "C" & destCellNum
    MarketNameCell = "D" & destCellNum
    MerchAreaCell = "E" & destCellNum
  Loop


End Sub
 
