Attribute VB_Name = "HTML"
Sub HTML()
    Application.DisplayAlerts = False
    If Range("A1").Value <> "Part Number" Then Range("A1").Value = "Part Number"
    ActiveSheet.UsedRange.Select
    Selection.WrapText = True
    Selection.WrapText = False
    On Error GoTo 27
    Rows("1:1").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.Delete Shift:=xlToLeft
27  Range("A1").Select
End Sub
