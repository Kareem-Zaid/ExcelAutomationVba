Attribute VB_Name = "Infineon"
Sub Infineon()
Attribute Infineon.VB_Description = "remove first row"
Attribute Infineon.VB_ProcData.VB_Invoke_Func = "i\n14"
'
' Infineon Macro
' remove first row
'
' Keyboard Shortcut: Ctrl+i
'
    'Insert "On Error Resume Next" at Ln 12 of ProcessFiles code
    If Range("A2").Value = "Product" Then
        Rows("1:1").Select
        Selection.Delete Shift:=xlUp
        ActiveSheet.UsedRange.Select
        Selection.WrapText = True
        Selection.WrapText = False
        Range("A1").Select
    End If
End Sub
