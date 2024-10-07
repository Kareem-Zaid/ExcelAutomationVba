Attribute VB_Name = "Sanken"
Sub Sanken()
    If Range("A1").Value = "Part Number" Then
        ActiveSheet.UsedRange.Select
        Selection.WrapText = True
        Selection.WrapText = False
        Range("A1").Select
    End If
End Sub
