;Start;
saveFile = Circ_PullList.docx
FileDelete %saveFile%

;Check for Template.docx
IfNotExist, Template.docx
{
	msgbox Cannot find Template.docx
	exit
}

;Get input file
FileSelectFile, xlsFile,,C:\Temp, Select File, *.xls*

;Check for input file or cancel to exit
If xlsFile =
{
	exit
}

;Open DOC file
template = %A_ScriptDir%\Template.docx
saveFilePath = %A_ScriptDir%\%saveFile%
wrd := ComObjCreate("Word.Application")
wrd.Visible := False

;Perform Mail Merge
doc := wrd.Documents.Open(template)
doc.MailMerge.MainDocumentType := 3 ;Mail merge type "directory"
doc.MailMerge.OpenDataSource(xlsFile,,,,,,,,,,,,,"SELECT * FROM [expiredHoldShelfRequestsList$] WHERE [Location] <> 'ILLIAD' AND [Location] <> 'Resource Sharing Long Loan'")
doc.MailMerge.Execute

;Apply banded row style
wrd.Selection.Tables(1).ApplyStyleRowBands := Not Selection.Tables(1).ApplyStyleRowBands
wrd.Selection.Tables(1).Style := "Grid Table 4"

;Add header row
wrd.Selection.InsertRowsAbove(1)
wrd.Selection.Tables(1).Rows(1).Height := 30
wrd.Selection.Cells.VerticalAlignment := 1
wrd.Selection.ParagraphFormat.Alignment := 1
wrd.Selection.Shading.BackgroundPatternColor := -587137025
wrd.Selection.Font.Italic := False
wrd.Selection.Font.Bold := True
wrd.Selection.TypeText("Requestor")
wrd.Selection.MoveRight(1)
wrd.Selection.TypeText("Title")
wrd.Selection.MoveRight(1)
wrd.Selection.TypeText("Barcode")
wrd.Selection.MoveRight(1)
wrd.Selection.TypeText("Held Since")
wrd.Selection.MoveRight(1)
wrd.Selection.TypeText("Held Until")
wrd.Selection.MoveRight(1)
wrd.Selection.TypeText("Location")
wrd.Selection.Rows.HeadingFormat := 9999998 ;Set header for each page

;Save and quit DOC file
wrd.ActiveDocument.SaveAs(saveFilePath)
wrd.DisplayAlerts := False
doc.Close
wrd.Quit

;Finish
IfNotExist, %saveFile%
{
	msgbox Cannot find %saveFile%
	exit
}
FileDelete %xlsFile%
run winword.exe %saveFile%