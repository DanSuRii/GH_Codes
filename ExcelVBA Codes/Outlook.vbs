Function BrowseForFolder(Optional OpenAt As Variant)
    
    Dim ShellApp As Object
    
    Set ShellApp = CreateObject("Shell.Application"). _
    BrowseForFolder(0, "Please choose a folder", 0, OpenAt)
    
    On Error Resume Next
    
    BrowseForFolder = ShellApp.self.Path
    
    On Error GoTo 0
    
    Set shellap = Nothing
    
End Function

Public Sub SaveAttachments()

    Dim objOL As Outlook.Application
    Dim objMsg As Outlook.MailItem
    Dim objAttachments As Outlook.Attachments
    Dim objSelection As Outlook.Selection
    
    Dim i As Long
    Dim lngCount As Long
    Dim strFile As String
    Dim strFolderPath As String
    
    
     Dim fldr As FileDialog
     Set fldr = FileDialog(msoFileDialogFolderPicker)
     With fldr
         .Title = "Select a folder"
         .AllowMultiSelect = False
         .InitialFileName = Application.DefaultFilePath
         If .Show <> -1 Then GoTo toExit
         strFolderPath = .SelectedItems(1)
     End With
    
    ' strFolderPath = CreateObject("Wscript.Shell").specialfolders(16)
    strFolderPath = BrowseForFolder("C:\")
    strFolderPath = fldr.SelectedItems.Item
    
    
    If strFolderPath = "" Then GoTo toExit
    If Dir(strFolderPath, vbDirectory) = "" Then GoTo toExit
    
    Set objSelection = Application.ActiveExplorer.Selection
    For Each objMsg In objSelection
        Set objAttachments = objMsg.Attachments
        For Each objAttachment In objAttachments
            strFile = objAttachment.FileName
            ' Debug.Print strFile
            strFile = strFolderPath & "\" & strFile
            ' Debug.Print strFile
            objAttachment.SaveAsFile strFile
        Next
    
    Next
    
toExit:
    Set objAttachments = Nothing
    Set objOL = Nothing
    
End Sub
