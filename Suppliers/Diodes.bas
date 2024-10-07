Attribute VB_Name = "Diodes"
Sub Diodes()
Attribute Diodes.VB_Description = "datasheet\npaste value\nremove hyperlink"
Attribute Diodes.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Diodes Macro
' datasheet paste value remove hyperlink
'

'
    If Range("A1").Value = "Part Number" Then
        Columns("B:B").Select
        Selection.Replace What:="=HYPERLINK(""", Replacement:="", LookAt:=xlPart _
            , SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
        Selection.Replace What:=""",""*", Replacement:="", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
        Cells.Select
        Range("C32").Activate
        Application.CutCopyMode = False
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        Selection.Hyperlinks.Delete
        ActiveSheet.UsedRange.Select
        Selection.WrapText = True
        Selection.WrapText = False
        Range("A1").Select
    End If
End Sub
