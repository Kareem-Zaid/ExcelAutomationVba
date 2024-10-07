Attribute VB_Name = "SKHynix"
Sub SK_Hynix()
Attribute SK_Hynix.VB_Description = "delete extra headers"
Attribute SK_Hynix.VB_ProcData.VB_Invoke_Func = " \n14"
'
' SK_Hynix Macro
' delete extra headers
'

'
    If Range("A3").Value = "Part No." Then
        Rows("1:2").EntireRow.Delete
        ActiveSheet.UsedRange.Select
        Selection.WrapText = True
        Selection.WrapText = False
        Range("A1").Select
    End If
End Sub
