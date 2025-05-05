This code prompts the user to search for emails with specific subject title and downloads all excel attachments within the past XX days (prompted by user)

```vba
Sub DownloadExcelFromOutlook()
  Dim OutlookApp As Object
  Dim OutlookNamespace As Object
  Dim Stores As Object
  Dim Store As Object
  Dim InboxFolder As Object, FolderStack As Collection
  Dim Folder As Object, SubFolder As Object
  Dim OutlookMail As Object
  Dim Attachment As Object
  Dim SubjectToFind As String
  Dim fd As FileDialog
  Dim SaveFolder As String
  Dim i As Long
  Dim DaysBack As Variant
  Dim ReceivedDateThreshold As Date
  Dim response As VbMsgBoxResult

  ' Prompt user for number of days
  DaysBack = InputBox("Enter number of days to look back for emails:", "Days Back")
  If Not IsNumeric(DaysBack) Or DaysBack <= 0 Then
    MsgBox "Invalid number of days.",vbExclamation
    Exit Sub
  End If

' Set the subject to search for
response = MsgBox("Look for email subject title starting with: 'COA Update:'?", vbYesNo + vbQuestion, "Confirmation")
If response = vbYes Then
  ' Execute below if the user clicks Yes
    SubjectToFind = "COA Update:"
Else
  ' Prompt for subject title if the user clicks No
  SubjectToFind = InputBox("Enter the subject or keyword to search for:", "Email Subject")
  If SubjectToFind = ""
Then MsgBox "Subject cannot be empty.", vbExclamation
    Exit Sub
  End If
End If

' Set the folder to save attachments
Set fd = Application.FileDialog(msoFileDialogFolderPicker)
fd.Title = "Please select the folder to save the excel attachments"

' Show the dialog box
If fd.Show = -1 Then
    ' If the user selects a folder, store the path
        SaveFolder = fd.SelectedItems(1)
        MsgBox "You selected: " & SaveFolder
Else ' If the user cancels the dialog box
  MsgBox "No folder selected. Exiting script..."
  Exit Sub
End If

' Calculate date threshold
ReceivedDateThreshold = Date - CLng(DaysBack)

' Initialize Outlook objects
Set OutlookApp = CreateObject("Outlook.Application")
Set OutlookNamespace = OutlookApp.GetNamespace("MAPI")
Set Stores = OutlookNamespace.Stores

' Loop through all accounts
For Each Store In Stores
  Set InboxFolder = Store.GetDefaultFolder(6) ' 6 = olFolderInbox

  ' Use a stack to process folders iteratively (avoids recursion)
  Set FolderStack = New Collection
  FolderStack.Add InboxFolder

  Do While FolderStack.Count > 0
    Set Folder = FolderStack(FolderStack.Count)
    FolderStack.Remove FolderStack.Count
    ' Check each item in folder

    For Each OutlookMail In Folder.Items
        If TypeName(OutlookMail) = "MailItem" Then
          If OutlookMail.ReceivedTime >= ReceivedDateThreshold Then
            If InStr(1, OutlookMail.Subject, SubjectToFind, vbTextCompare) > 0 Then
              For Each Attachment In OutlookMail.Attachments
                If LCase(Right(Attachment.Filename, 5)) = ".xlsx" Or _ LCase(Right(Attachment.Filename, 4)) = ".xls" Then
                  Attachment.SaveAsFile SaveFolder & "\" & Attachment.Filename
                  TotalDownloaded = TotalDownloaded + 1
                End If
              Next Attachment
          End If
        End If
      End If
    Next OutlookMail

    ' Add subfolders to the stack
        For Each SubFolder In Folder.Folders
          FolderStack.Add SubFolder
        Next SubFolder
  Loop
Next Store

MsgBox TotalDownloaded & " Excel attachment(s) downloaded from all inboxes and subfolders.", vbInformation
End Sub
```
