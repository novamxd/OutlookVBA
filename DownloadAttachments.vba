Dim xFso As FileSystemObject

''' <summary>
''' The main method of this module. Downloads all the attachments of all the selected emails. Tested with >600 emails and >600 attachments.
''' Does the downloading in batches of 50 emails at a time to prevent server and memory limitations
''' </summary>
''' <returns></returns>
Public Sub SaveAttachments()

    'create our global FSO object to save on allocations
    Set xFso = CreateObject("Scripting.FileSystemObject")
    
    'get the user's documents folder
    Dim xFolderPath As String
    xFolderPath = CreateObject("WScript.Shell").SpecialFolders(16)
    xFolderPath = xFolderPath & "\Attachments\"
    
    'make the attachments folder if it doesn't already exist
    If VBA.Dir(xFolderPath, vbDirectory) = vbNullString Then
        VBA.MkDir xFolderPath
    End If
    
    'get our selection so we can get the count for our batches
    Dim xSelection As Outlook.Selection
    Set xSelection = Outlook.Application.ActiveExplorer.Selection
    
    Dim xStart As Long
    xStart = 1
    
    Dim xEnd As Long
    xEnd = xStart
    
    'process our email in batches to prevent memory leaks (too many items open)
    While xEnd < xSelection.Count
    
        'assume batch of 50 emails
        xEnd = xStart + 50 - 1
        
        'crop the end if needed
        If xEnd > xSelection.Count Then
            xEnd = xSelection.Count
        End If
    
        'process this batch of 50 emails
        Call ProcessEmailBatch(xFolderPath, xStart, xEnd)
        
        'move our cursor along
        xStart = xEnd + 1
        
    Wend

    'clear our global reference
    Set xFso = Nothing

    'let the user know
    MsgBox "All Done!", vbOKOnly, "DownloadAttachments.vba"

End Sub

''' <summary>
''' Downloads all the attachments from the emails in a range of the selection. This method exists because in most
''' cases Exchange will not allow more than 250 emails open at a time. The only way to ensure the emails are closed
''' are to properly dispose of them in VBA, and you do that by leaving scope. The selection is what maintains the
''' connection, so we need a new connection per batch.
''' </summary>
''' <param name="xFolderPath">The destination folder where the attachment should end up</param>
''' <param name="xStart">The starting position in the selected range</param>
''' <param name="xEnd">The end position in the selected range</param>
''' <returns></returns>
Sub ProcessEmailBatch(xFolderPath As String, xStart As Long, xEnd As Long)
    
    'get our selection (what's held open on the server)
    Dim xSelection As Outlook.Selection
    Set xSelection = Outlook.Application.ActiveExplorer.Selection
    
    'process our email items
    Dim xCurrent As Long
    For xCurrent = xStart To xEnd
        Call ProcessEmail(xFolderPath, xSelection.Item(xCurrent))
    Next

End Sub

''' <summary>
''' Downloads all the attachments in the provided email to the destination folder. Renames attachments if there are existing.
''' </summary>
''' <param name="xFolderPath">The destination folder where the attachment should end up</param>
''' <param name="xMailItem">The target email that needs the attachments downloaded</param>
''' <returns></returns>
Sub ProcessEmail(xFolderPath As String, xMailItem As MailItem)
    
    Dim xAttachments As Outlook.Attachments
    Set xAttachments = xMailItem.Attachments
    
    Dim xAttCount As Long
    xAttCount = xAttachments.Count
    
    'stop if there are no attachments
    If xAttCount = 0 Then
        Exit Sub
    End If
        
    'save each of our attachments to disk
    Dim xCurrent As Long
    For xCurrent = xAttCount To 1 Step -1
        Call ProcessAttachment(xFolderPath, xAttachments.Item(xCurrent))
    Next
    
End Sub

''' <summary>
''' Downloads the attachment to the folder. Renames the attachment as needed before resting on a file name
''' </summary>
''' <param name="xFolderPath">The destination folder where the attachment should end up</param>
''' <param name="xAttachment">The target attachment from an email</param>
''' <returns></returns>
Sub ProcessAttachment(xFolderPath As String, xAttachment As Attachment)
    
    'stop processing if it's an embedded attachment
    If IsEmbeddedAttachment(xAttachment) = True Then
        Exit Sub
    End If
        
    'get our unique file name that doesn't already exist
    Dim xFilePath As String
    xFilePath = GetUniqueFilePath(xFolderPath, xAttachment.FileName)
    
    'save the attachment to disc
    xAttachment.SaveAsFile xFilePath

End Sub

''' <summary>
''' Attempts to generate a unique name for the given file name by incrementing a counter and appending it to the file
''' </summary>
''' <param name="xFolderPath">The destination folder where the file should end up</param>
''' <param name="xFileName">The original name of the file</param>
''' <returns>The complete path to the file</returns>
Function GetUniqueFilePath(xFolderPath As String, xFileName As String) As String
    
    Dim xTry As Boolean
    xTry = True
    
    Dim xName As String
    xName = xFso.GetBaseName(xFileName)
    
    Dim xExtension As String
    xExtension = xFso.GetExtensionName(xFileName)
    
    'start with our default path
    Dim xPath As String
    xPath = xFolderPath & "\" & xName & "." + xExtension
        
    Dim xCount As Integer
    xCount = 0
    
    'keep trying until we're g2g
    While xTry
    
        If xFso.FileExists(xPath) Then
            xCount = xCount + 1
            xPath = xFso.BuildPath(xFolderPath, xName & " " & xCount & "." + xExtension)
        Else
            xTry = False
        End If
    
    Wend
    
    GetUniqueFilePath = xPath

End Function

''' <summary>
''' Checks whether the provided attachment is an embedded attachment or a separate file within the email
''' </summary>
''' <param name="xAttach">The attachment to evaluate</param>
''' <returns>Whether the attachment is embedded or not</returns>
Function IsEmbeddedAttachment(xAttach As Attachment)
    Dim xItem As MailItem
    Dim xCid As String
    Dim xID As String
    Dim xHtml As String
    IsEmbeddedAttachment = False
    Set xItem = xAttach.Parent
    If xItem.BodyFormat <> olFormatHTML Then Exit Function
    xCid = ""
    xCid = xAttach.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F")
    If xCid <> "" Then
        xHtml = xItem.HTMLBody
        xID = "cid:" & xCid
        If InStr(xHtml, xID) > 0 Then
            IsEmbeddedAttachment = True
        End If
    End If
End Function
