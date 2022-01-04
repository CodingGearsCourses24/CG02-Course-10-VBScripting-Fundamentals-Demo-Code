'------------------------------------------------------------
' Script Overview:
'	- Create a excel document
'	- Add two new sheets - MyTemp1
'	- Delete the default sheet "sheet1"
'	- Add some data/text
'	- Format the text (Font - Type, Size & Color)
'------------------------------------------------------------

Option Explicit

Dim objExcel
Dim objWorkbook
Dim objWorkSheet
Dim ObjSheet

Dim sheetName
Dim strDirectory
Dim excelDocName

strDirectory = "D:\GitBucket\vbscripting-course\02-indev\excel\tmp_1"
excelDocName = "MyExcelDocument3.xlsx"

'Excel Object
Set objExcel = CreateObject("Excel.Application")

'Excel Document Visible Property
objExcel.visible = False

'Workbook Object 
Set objWorkbook = objExcel.Workbooks.Add 

'Add a new sheet with a name & adding some text [Cells(Row,Column)]
objWorkbook.Sheets.Add.Name = "MyTemp1"
objWorkbook.Sheets("MyTemp1").Cells(1,1).value = "I am working with Excel using VBScript" 
objWorkbook.Sheets("MyTemp1").Range("A2").value = "It is easy!" 

'Formatting - A1
objWorkbook.Sheets("MyTemp1").Range("A1").Font.Name = "Cambria"
objWorkbook.Sheets("MyTemp1").Range("A1").Font.Bold = True
objWorkbook.Sheets("MyTemp1").Range("A1").Font.Italic = True
objWorkbook.Sheets("MyTemp1").Range("A1").Font.Size = 18
objWorkbook.Sheets("MyTemp1").Range("A1").Font.ColorIndex = 5 'Blue

'Formatting - A2
objWorkbook.Sheets("MyTemp1").Range("A2").Font.Name = "Cambria"
objWorkbook.Sheets("MyTemp1").Range("A2").Font.Bold = True
objWorkbook.Sheets("MyTemp1").Range("A2").Font.Italic = True
objWorkbook.Sheets("MyTemp1").Range("A2").Font.Size = 18
objWorkbook.Sheets("MyTemp1").Range("A2").Font.ColorIndex = 10 'Green

'Delete the default sheet1 
objExcel.Sheets("sheet1").Delete

'Save, Close & Quit
objWorkbook.SaveAs strDirectory & "\" &  excelDocName
objWorkbook.Close 
objExcel.Quit

MsgBox "Done! 					" , 0, "Status"

'Release Memory
Set objExcel = Nothing
Set objWorkbook = Nothing