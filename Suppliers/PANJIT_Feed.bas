Attribute VB_Name = "PANJIT_Feed"
Sub PANJIT_Feed()
Attribute PANJIT_Feed.VB_Description = "merge row7&8\ndelete extra headers\nremove hyperlinks"
Attribute PANJIT_Feed.VB_ProcData.VB_Invoke_Func = "P\n14"
'
' PANJIT_Feed Macro
' delete extra header remove Hyperlinks header joiner for 1st 27 cells
'
' Keyboard Shortcut: Ctrl+Shift+P
'
    Application.DisplayAlerts = False
    If Range("A2").Value = "Part Number" And Range("B1").Value = "" And Range("AB2").Value = "" Then
        Cells.Select
        Selection.Hyperlinks.Delete
        Rows("1:1").Select
        Selection.Delete Shift:=xlUp
        Rows("1:1").Select
        Selection.Insert Shift:=xlDown
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
