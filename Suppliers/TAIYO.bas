Attribute VB_Name = "TAIYO"
Sub TAIYO()
Attribute TAIYO.VB_Description = "remove headers\nedit formulas"
Attribute TAIYO.VB_ProcData.VB_Invoke_Func = " \n14"
'
' TAIYO Macro
' rename header cell delete extra headers edit formulas
'

'
    If Range("C3").Value = "Shape Symbol" Then Range("C3").Value = "Part Number"
    If Range("C3").Value = "Part Number" Then
        Rows("1:2").Select
        Selection.Delete Shift:=xlUp
        Range("A1").Select
        On Error Resume Next
        Selection.SpecialCells(xlCellTypeFormulas, 23).Select
        Selection.Replace What:="=", Replacement:="+", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
        Selection.Replace What:="/", Replacement:="/-", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
        ActiveSheet.UsedRange.Select
        Selection.WrapText = True
        Selection.WrapText = False
        Range("A1").Select
    End If
End Sub
