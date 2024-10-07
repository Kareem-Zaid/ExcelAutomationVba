Attribute VB_Name = "Nexperia"
Sub Nexperia_dh()
Attribute Nexperia_dh.VB_Description = "delete extra headers\nnexperia"
Attribute Nexperia_dh.VB_ProcData.VB_Invoke_Func = "n\n14"
'
' Nexperia_dh Macro
' paste value delete extra headers delete shape
'
' Keyboard Shortcut: Ctrl+n
'
    If Range("A10").Value = "Type number" Then
        Range("a1", Range("a1").End(xlDown).End(xlToRight)).Select
        Selection.Copy
        ActiveWindow.ScrollColumn = 1
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        Rows("1:9").Select
        Selection.Delete Shift:=xlUp
        ActiveSheet.Shapes.Range(Array("Picture 1")).Select
        Selection.Delete
        ActiveSheet.UsedRange.Select
        Selection.WrapText = True
        Selection.WrapText = False
        Range("A1").Select
    End If
End Sub
