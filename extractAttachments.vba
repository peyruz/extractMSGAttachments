Dim strAttachmentFolder As String

Sub ExtractAttachmentsFromEmailsStoredinWindowsFolder()
    Dim objShell, objWindowsFolder As Object
 
    'Select a Windows folder
    Set objShell = CreateObject("Shell.Application")
    Set objWindowsFolder = objShell.BrowseForFolder(0, "Select a Windows Folder:", 0, "")
 
    If Not objWindowsFolder Is Nothing Then
       'Create a new folder for saving extracted attachments
       strAttachmentFolder = "<------ Path to the output folder ------>"
       'MkDir (strAttachmentFolder)
       Call ProcessFolders(objWindowsFolder.self.Path & "\")
       MsgBox "Completed!", vbInformation + vbOKOnly
    End If
End Sub

Sub ProcessFolders(strFolderPath As String)
    Dim objFileSystem As Object
    Dim objFolder As Object
    Dim objFiles As Object
    Dim objFile As Object
    Dim objItem As Object
    Dim i As Long
    Dim objSubFolder As Object
    Dim fileName As String
    
    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFileSystem.GetFolder(strFolderPath)
    Set objFiles = objFolder.Files
    
    o = 1
    For Each objFile In objFiles
        Debug.Print (o)
        If objFileSystem.GetExtensionName(objFile) = "msg" Then
           'Open the Outlook emails stored in Windows folder
           Set objItem = Session.OpenSharedItem(objFile.Path)

           If TypeName(objItem) = "MailItem" Then
              If objItem.Attachments.Count > 0 Then
                 'Extract attachments
                 For i = objItem.Attachments.Count To 1 Step -1
                    If InStr(objItem.Attachments(i).fileName, "xls") Then
                    fileName = "file" & CStr(o) & " " & Month(objItem.ReceivedTime) & "-" & _
                        Day(objItem.ReceivedTime) & "-" & Year(objItem.ReceivedTime) & "_" & _
                        Hour(objItem.ReceivedTime) & Minute(objItem.ReceivedTime) & ".xlsx"
                    
                    objItem.Attachments(i).SaveAsFile strAttachmentFolder & fileName
                    Debug.Print (objItem.Attachments(i).fileName & " saved as " & fileName)
                    o = o + 1
                    End If
                 Next
              End If
           End If
        End If
    Next
 
    'Process all subfolders recursively
    If objFolder.SubFolders.Count > 0 Then
       For Each objSubFolder In objFolder.SubFolders
           If ((objSubFolder.Attributes And 2) = 0) And ((objSubFolder.Attributes And 4) = 0) Then
              Call ProcessFolders(objSubFolder.Path)
           End If
       Next
    End If
End Sub
