Attribute VB_Name = "ONSemi"
Sub ONSemi()
Attribute ONSemi.VB_Description = "extract datasheet url\npaste value"
Attribute ONSemi.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ONSemi Macro
' extract datasheet url paste value
'

'
    If Range("A1").Value = "Product" Then
        Columns("B:B").Select
        Selection.Replace What:="=HYPERLINK(""", Replacement:="", LookAt:=xlPart _
            , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
        Selection.Replace What:=""",""*", Replacement:="", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
        Cells.Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        ActiveSheet.UsedRange.Select
        Selection.WrapText = True
        Selection.WrapText = False
        Range("A1").Select
    End If
End Sub
