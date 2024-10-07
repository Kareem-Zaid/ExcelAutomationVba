Attribute VB_Name = "CSVtoXLSX"
Sub CSVtoXLSX()

    Dim myfile As String
    Dim oldfname As String, newfname As String
    Dim workfile
    Dim FolderName As FileDialog
    Dim folderNameItem As Variant


    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

'   Capture name of current file
    myfile = ActiveWorkbook.Name

'   Set folder name to work through
' DEL    folderName = "C:\Users\m\Desktop\CSVtoEXCEL\"
    Set FolderName = Application.FileDialog(msoFileDialogFolderPicker)
    If FolderName.Show = -1 Then
        folderNameItem = FolderName.SelectedItems(1) & Application.PathSeparator

'   Loop through all CSV files in folder
    workfile = Dir(folderNameItem & "*.CSV")
    Do While workfile <> ""
'       Open CSV file
        Workbooks.Open Filename:=folderNameItem & workfile
'       Capture name of old CSV file
        oldfname = ActiveWorkbook.FullName
'       Convert to XLSX
        newfname = folderNameItem & Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 4) & ".xlsx"
        ActiveWorkbook.SaveAs Filename:=newfname, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        ActiveWorkbook.Close
'       Delete old CSV file
        Kill oldfname
        Windows(myfile).Activate
        workfile = Dir()
    Loop

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    End If
End Sub
