Attribute VB_Name = "Analog"
Sub Analog()
Attribute Analog.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Analog Macro
' delete extra sheets paste value header joiner for 1st 30 cells
'

'
    If Sheets("Raw Data Display").Range("A1").Value = "Part Number" And Sheets("Raw Data Display").Range("AE1").Value = "" Then
        Application.DisplayAlerts = False
        Sheets(Array("Cover", "Web Display")).Delete
        Cells.Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Rows("1:1").Select
        Application.CutCopyMode = False
        Selection.Insert Shift:=xlDown
        Range("A1").Select
        ActiveCell.FormulaR1C1 = _
            "=IF(AND(R[1]C="""",R[2]C=""""),""K&S"",IF(R[2]C<>"""",R[1]C&"" ""&R[2]C,R[1]C))"
        Range("AD1").Select
        Range(Selection, Selection.End(xlToLeft)).Select
        Selection.FillRight
        Selection.Copy
        ActiveSheet.Paste
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Selection.Replace What:="K&S", Replacement:="", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        Range("A1").Find What:="", After:=ActiveCell, LookIn:=xlFormulas, _
              LookAt:=xlPart, SearchOrder:=xlByRows, _
              SearchDirection:=xlNext, MatchCase:=False, _
              SearchFormat:=False
        Rows("2:3").Select
        Application.CutCopyMode = False
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
