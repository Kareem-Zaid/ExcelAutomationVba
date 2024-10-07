Attribute VB_Name = "IDT"
Sub IDT()
Attribute IDT.VB_Description = "rmv 3 rows\nremove hyperlinks"
Attribute IDT.VB_ProcData.VB_Invoke_Func = " \n14"
'
' IDT Macro
' rmv 3 rows remove hyperlinks
'

'
    If Range("A4").Value = "Product ID" Then
        Rows("1:3").Select
        Selection.Delete Shift:=xlUp
        Cells.Select
        Selection.Hyperlinks.Delete
        ActiveSheet.UsedRange.Select
        Selection.WrapText = True
        Selection.WrapText = False
        Range("A1").Select
    End If
End Sub
