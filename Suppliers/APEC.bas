Attribute VB_Name = "APEC"
Sub APEC()
'
' Macro9 Macro
' paste value delete extra rows&columns first 2 columns unmerge-fill header joiner for first 35 cells rplc ""Ctrl+J"" by "" "" in header
'

'
    If Range("C2").Value = "Part Number" Or Range("B2").Value = "Part Number" And Range("AJ2").Value = "" Then
        If Range("C2").Value = "Part Number" Then
            Columns("C:D").Select
            Range("C2").Activate
            Application.CutCopyMode = False
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Application.CutCopyMode = False
            Selection.Copy
            ActiveSheet.Paste
            Rows("1:1").Select
            Range("C1").Activate
            Application.CutCopyMode = False
            Selection.Delete Shift:=xlUp
            Columns("A:B").Select
            Selection.Delete Shift:=xlToLeft
        ElseIf Range("B2").Value = "Part Number" Then
            Columns("B:C").Select
            Range("B2").Activate
            Application.CutCopyMode = False
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Application.CutCopyMode = False
            Selection.Copy
            ActiveSheet.Paste
            Rows("1:1").Select
            Range("B1").Activate
            Application.CutCopyMode = False
            Selection.Delete Shift:=xlUp
            Columns("A").Select
            Selection.Delete Shift:=xlToLeft
        End If
        Rows("1:2").Select
        Selection.Copy
        ActiveSheet.Paste
        Range("A1:AI1").Select
        Selection.SpecialCells(xlCellTypeBlanks).Select
        Application.CutCopyMode = False
        Selection.FormulaR1C1 = "=IF(R[1]C="""",""K&S"",RC[-1])"
        Rows("1:2").Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Selection.Replace What:="K&S", Replacement:="", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        Columns("A:B").Select
        Selection.Copy
        Selection.SpecialCells(xlCellTypeBlanks).Select
        Selection.FormulaR1C1 = "=IF(R[1]C="""",""K&S"",R[-1]C)"
        Columns("A:B").Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Selection.Replace What:="K&S", Replacement:="", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Selection.Replace What:="0", Replacement:="", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        Range("A1").Select
        Selection.EntireRow.Insert , CopyOrigin:=xlFormatFromRightOrBelow
        ActiveCell.FormulaR1C1 = _
            "=IF(AND(R[1]C="""",R[2]C="""",R[1]C=R[2]C),""K&S"",IF(AND(R[2]C<>"""",R[1]C<>R[2]C),R[1]C&"" ""&R[2]C,R[1]C))"
        Range("AI1").Select
        Range(Selection, Selection.End(xlToLeft)).Select
        Selection.FillRight
        Selection.Copy
        ActiveSheet.Paste
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Selection.Replace What:="K&S", Replacement:="", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        Selection.Replace What:="" & Chr(10) & "", Replacement:=" ", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
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
