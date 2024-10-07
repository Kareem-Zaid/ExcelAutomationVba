Attribute VB_Name = "Torex"
Sub Torex()
Attribute Torex.VB_Description = "replace line breaks in header by space\nclean header"
Attribute Torex.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Torex Macro
' replace line breaks in header by space clean header for first 30 cells
'

'
    Application.DisplayAlerts = False
    If Range("A1").Value = "Series Name" And Range("AE1").Value = "" Then
        Rows("1:1").Select
        Selection.Replace What:="" & Chr(10) & "", Replacement:=" ", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
' clean header for first 30 cells
        Selection.Insert Shift:=xlDown
        Range("A1").Select
        ActiveCell.FormulaR1C1 = "=IF(R[1]C="""",""K&S"",CLEAN(R[1]C))"
        Range("AD1").Select
        Range(Selection, Selection.End(xlToLeft)).Select
        Selection.FillRight
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Selection.Replace What:="K&S", Replacement:="", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        Range("A1").Find What:="", After:=ActiveCell, LookIn:=xlFormulas, _
              LookAt:=xlPart, SearchOrder:=xlByRows, _
              SearchDirection:=xlNext, MatchCase:=False, _
              SearchFormat:=False
        Rows("2:2").Select
        Selection.Delete Shift:=xlUp
        ActiveSheet.UsedRange.Select
        Selection.WrapText = True
        Selection.WrapText = False
        On Error GoTo 27
        Rows("1:1").Select
        Selection.SpecialCells(xlCellTypeBlanks).Select
        Selection.Delete Shift:=xlToLeft
27        Range("A1").Select
    End If
End Sub
