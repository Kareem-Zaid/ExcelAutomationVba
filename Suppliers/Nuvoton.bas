Attribute VB_Name = "Nuvoton"
Sub Nuvoton()
' kaze first non-recorded-by-macro code
' all to "Part No." delete shape delete extra headers
    If Range("A9").Value = "Part No." Or Range("A9").Value = "Part No" Or Range("A9").Value = "Part number" Then
        Range("A9").Value = "Part No."
        ActiveSheet.Shapes.Range(Array("Picture 1")).Delete
        Rows("1:8").EntireRow.Delete
        ActiveSheet.UsedRange.Select
        Selection.WrapText = True
        Selection.WrapText = False
        Range("A1").Select
    End If
End Sub
