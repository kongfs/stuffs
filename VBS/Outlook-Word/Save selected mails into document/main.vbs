Const MAILS_FILE_NAME = "emails.docx"


Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim currentDir : currentDir = fso.GetParentFolderName(WScript.ScriptFullName) & "/"

Dim emailsFile : emailsFile = currentDir & MAILS_FILE_NAME

Dim outlookApp : Set outlookApp = GetObject("", "Outlook.Application")

Dim wordApp : Set wordApp = CreateObject("Word.Application")
wordApp.Visible = True

Dim emailsDoc : Set emailsDoc = wordApp.Documents.Add

For Each selectedItem In outlookApp.ActiveExplorer.Selection
	If selectedItem.Class = 43 Then 'Outlook.olMail
	
		emailsDoc.Range.Text = vbNullString
	
		selectedItem.GetInspector().WordEditor.Range.FormattedText.Copy
		
		Dim oRange : Set oRange = emailsDoc.Content
		oRange.Collapse 0  'Word.WdCollapseDirection.wdCollapseEnd
		oRange.PasteAndFormat 22 'Word.wdFormatPlainText

	End If
Next

emailsDoc.SaveAs2 emailsFile
emailsDoc.Close False 
wordApp.Quit

MsgBox "Completed"

