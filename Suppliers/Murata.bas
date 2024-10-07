Attribute VB_Name = "Murata"
Sub Murata()
' delete extra headers tricky data (post-header row)
    If Range("A6").Value = "Part Number" Or Range("A7").Value = "Part Number" Or Range("A7").Value = "IC Part Number" Or Range("A1").Value = "Part Number" Then
        If Range("A6").Value = "Part Number" Then
            If Range("A7").Value = "Non-Preferred" And Range("H6").Value = "" Then
                Range("G6").Value = "Preferred/Non-Preferred"
                Range("H6").Value = "Input Power/Allowable Power(%)"
                Rows(7).EntireRow.Delete
            End If
            Rows("1:5").EntireRow.Delete
        ElseIf Range("A7").Value = "Part Number" Or Range("A7").Value = "IC Part Number" Then Rows("1:6").EntireRow.Delete
        End If
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
