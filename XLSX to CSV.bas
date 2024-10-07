Attribute VB_Name = "XLSXtoCSV"
Sub XLSXtoCSV()
Attribute XLSXtoCSV.VB_ProcData.VB_Invoke_Func = " \n14"

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

'   Loop through all XLSX files in folder
    workfile = Dir(folderNameItem & "*.XLSX")
    Do While workfile <> ""
'       Open XLSX file
        Workbooks.Open Filename:=folderNameItem & workfile
'       Capture name of old XLSX file
        oldfname = ActiveWorkbook.FullName
'       Convert to CSV
        newfname = folderNameItem & Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5) & ".csv"
        ActiveWorkbook.SaveAs Filename:=newfname, FileFormat:=xlCSV, CreateBackup:=False
        ActiveWorkbook.Close
'       Delete old XLSX file
        Kill oldfname
        Windows(myfile).Activate
        workfile = Dir()
    Loop

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    End If
End Sub

