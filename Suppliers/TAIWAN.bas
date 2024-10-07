Attribute VB_Name = "Taiwan"
Sub Taiwan()
Attribute Taiwan.VB_Description = "delete extra headers & footers"
Attribute Taiwan.VB_ProcData.VB_Invoke_Func = " \n14"
'
' TAIWAN Macro
' delete extra headers & footers
'

'
    If Range("A7").Value = "Part number" Then
        Rows("1:6").Select
        Selection.Delete Shift:=xlUp
        ActiveSheet.Cells(Rows.Count, "A").End(xlUp).EntireRow.Delete
        Rows("1:1").Select
        Selection.Replace What:="" & Chr(10) & "", Replacement:=" ", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
        ActiveSheet.UsedRange.Select
        Selection.WrapText = True
        Selection.WrapText = False
        Range("A1").Select
    End If
End Sub
