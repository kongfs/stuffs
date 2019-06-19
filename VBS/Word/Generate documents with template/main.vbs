Main()

Const PREFIX = "["
Const SUFFIX = "]"

Const DATA_FILE_NAME = "data.xlsx"
Const HEADER_ROW = 1
Const TEMPLATE_FILE = "template.docx"
Const OUTPUT_DIR_NAME = "output"

Const FILE_SYSTEM_OBJECT = "Scripting.FileSystemObject"
Const EXCEL_APPLICATION = "Excel.Application"
Const WORD_APPLICATION = "Word.Application"

Public Sub Main()

	Dim fso : Set fso = CreateObject(FILE_SYSTEM_OBJECT)
	
	Dim currentDir : currentDir = fso.GetParentFolderName(WScript.ScriptFullName) & "/"
	
	Dim templateFile : templateFile = currentDir & TEMPLATE_FILE
	
	Dim outputDir : outputDir = currentDir & OUTPUT_DIR_NAME & "/"
 
	If fso.FolderExists(outputDir) = False Then fso.CreateFolder outputDir

	Dim dataFile : dataFile = currentDir & DATA_FILE_NAME
	
	Dim excelApp : Set excelApp = CreateObject(EXCEL_APPLICATION)
	excelApp.Visible = True
	  
	Dim dataBook : Set dataBook = excelApp.WorkBooks.Open(dataFile)
	
	Dim dataSheet : Set dataSheet = dataBook.Sheets(1)
	 
	Dim wordApp : Set wordApp = CreateObject(WORD_APPLICATION)
	wordApp.Visible = True
 
    For row = HEADER_ROW + 1 To dataSheet.UsedRange.Rows.Count
     
		If dataSheet.Cells(row, 1).Value = Empty Then Exit For
		
		Dim templateDoc : Set templateDoc = wordApp.Documents.Open(templateFile)
	 
		For col = 1 To dataSheet.UsedRange.Columns.Count
		
			If dataSheet.Cells(HEADER_ROW, col).Value = Empty Then Exit For

			Dim findText : findText = PREFIX & dataSheet.Cells(HEADER_ROW, col).Value & SUFFIX 
			Dim replaceText : replaceText = dataSheet.Cells(row, col).Value 

			With templateDoc.Range.Find
				.ClearFormatting
				.Replacement.ClearFormatting 
				.Text = findText
                .Replacement.Text = replaceText
                .Forward = True
				.Execute ,,,,,,,,,,2 'wdReplaceAll
			End With
 
		Next
        
		Dim newFile : newFile = outputDir &  row & ".docx"
		
        templateDoc.SaveAs2(newFile)
        templateDoc.Close(False)
    
    Next
	
    wordApp.Quit
	
	dataBook.Close False
    excelApp.Quit
	
	MsgBox "Completed"
	
End Sub

