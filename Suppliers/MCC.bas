Attribute VB_Name = "MCC"
Sub MCC()
Attribute MCC.VB_Description = "delete extra headers"
Attribute MCC.VB_ProcData.VB_Invoke_Func = "Q\n14"
'
' MCCSemi Macro
' delete extra headers
'
' Keyboard Shortcut: Ctrl+Shift+Q
'
    If Range("A2").Value = "Product" Or UCase(Range("C1").Value) = "PNUMBER" Or UCase(Range("B1").Value) = "PNUMBER" Or Range("A2").Value = "Part Number" Then
        If Range("A2").Value = "Product" And Range("A3").Value = "" Then
            ActiveSheet.Shapes.Range(Array("MCC-Logo")).Select
            Selection.Delete
            Range("1:1,3:3").Select
            Range("A3").Activate
            Selection.Delete Shift:=xlUp
            Rows("1:1").Replace What:="" & Chr(10) & "", Replacement:=" ", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
        ElseIf UCase(Range("C1").Value) = "PNUMBER" Then Range("C1").Value = "Product"
        ElseIf UCase(Range("B1").Value) = "PNUMBER" Then Range("B1").Value = "Product"
        ElseIf Range("A2").Value = "Part Number" And Range("Z2").Value = "" Then
            Range("A2").Value = "Product"
            ActiveSheet.Shapes.Range(Array("MCC-Logo")).Delete
            Rows("1:1").Delete
            Rows("1:1").Select
            Application.CutCopyMode = False
            Selection.Insert Shift:=xlDown
            Range("A1").Select
        ActiveCell.FormulaR1C1 = _
            "=IF(AND(R[1]C="""",R[2]C="""",R[1]C=R[2]C),""K&S"",IF(AND(R[2]C<>"""",R[1]C<>R[2]C),R[1]C&"" ""&R[2]C,R[1]C))"
            Range("Y1").Select
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
        End If
        ActiveSheet.UsedRange.Select
        Selection.WrapText = True
        Selection.WrapText = False
        Range("A1").Select
    End If
End Sub
