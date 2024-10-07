Attribute VB_Name = "Maxim"
Sub Maxim()
Attribute Maxim.VB_Description = "delete extra headers\nconditional additional row remover\nremove hyperlinks"
Attribute Maxim.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Maxim Macro
' remove hyperlinks conditional additional row remover delete extra headers
'

'
    If Range("A6").Value = "Part Number" Or Range("A5").Value = "Part Number" Then
        If Range("A7").Value = "Current Selections:" Then Rows(7).EntireRow.Delete
        If Range("A6").Value = "Current Selections:" Then Rows(6).EntireRow.Delete
        Cells.Select
        Selection.Hyperlinks.Delete
        If Range("A6").Value = "Part Number" Then
            Rows("1:5").Select
            Selection.Delete Shift:=xlUp
        ElseIf Range("A5").Value = "Part Number" Then
            Rows("1:4").Select
            Selection.Delete Shift:=xlUp
        End If
        ActiveSheet.UsedRange.Select
        Selection.WrapText = True
        Selection.WrapText = False
        Range("A1").Select
    End If
End Sub
