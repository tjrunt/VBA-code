Attribute VB_Name = "slangCode"
Sub Slang()
Attribute Slang.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Slang Macro
'

'
    Selection.Replace What:="SLANG", Replacement:="TB", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False
    ActiveWindow.ScrollWorkbookTabs Position:=xlFirst
    Sheets("PivotTable").Select
    ActiveWorkbook.RefreshAll
    Call piviot_format
    Call reviseTabs_PageNameMarket
End Sub
