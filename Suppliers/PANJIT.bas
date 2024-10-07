Attribute VB_Name = "PANJIT"
Sub PANJIT()
Attribute PANJIT.VB_Description = "merge row7&8\ndelete extra headers\nremove hyperlinks"
Attribute PANJIT.VB_ProcData.VB_Invoke_Func = "P\n14"
'
' PANJIT Macro
' delete extra headers remove Hyperlinks header joiner for 1st 27 cells
'
' Keyboard Shortcut: Ctrl+Shift+P
'
    If Range("A7").Value = "Part Number" And Range("A6").Value = "" And Range("AB7").Value = "" Then
        Cells.Select
        Selection.Hyperlinks.Delete
        Rows("1:5").Select
        Selection.Delete Shift:=xlUp
        Range("A1").Select
        ActiveCell.FormulaR1C1 = _
            "=IF(AND(R[1]C="""",R[2]C=""""),""K&S"",IF(R[2]C<>"""",R[1]C&"" ""&R[2]C,R[1]C))"
        Range("AA1").Select
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
