WorkingDir = ""
Extension = ".xlsx"

Dim fso, myFolder, fileColl, aFile, FileName, SaveName
Dim objExcel, objWorkbook

Set WshShell = CreateObject("WScript.Shell")

Set fso = CreateObject("Scripting.FilesystemObject")
'Set myFolder = fso.GetFolder(WorkingDir)
'Set myFolder = fso.GetAbsolutePathName(".")
WorkingDir = WshShell.CurrentDirectory
Set myFolder = fso.GetFolder(WorkingDir)
'fileColl = myFolder.Files

SEt objExcel = CreateObject("Excel.Application") 

For Each aFile In myFolder.Files
    ext = right(aFile.Name, 5)
    If UCase(ext) = UCase(extension) Then
    
        'Open excel
        FileName = Left(aFile, InStrRev(aFile, "."))
        Set objWorkbook = objExcel.Workbooks.Open(aFile)
        SaveName = FileName & "csv"
        objWorkbook.SaveAs SaveName, 6
        objWorkbook.Close

    End If
Next

Set objWorkbook = Nothing
Set objExcel = Nothing
Set fso = Nothing
Set myFolder = Nothing
Set fileColl = Nothing