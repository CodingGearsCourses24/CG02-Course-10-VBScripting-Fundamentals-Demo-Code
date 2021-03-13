'------------------------------------------------------------
' We will Explore:
'	- [Excel.Application] -> [Workbook] -> [Sheets] -> [Rows, Columns, Cells]
'	- OPEN an existing Microsoft Excel Document
'	- Delete a sheet (MyTemp2)
'	- Save the document.
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
excelDocName = "MyExcelDocument1.xlsx"

'Excel Object
Set objExcel = CreateObject("Excel.Application")
objExcel.DisplayAlerts = False  

'Excel Document Visible Property
objExcel.visible = False

'Getting an workbook object (using specific file)
Set objWorkbook = objExcel.Workbooks.Open(strDirectory & "\" &  excelDocName) 

'Delete a sheet
objWorkbook.Sheets("MyTemp2").Select
objWorkbook.Sheets("MyTemp2").Delete

'Save, Close & Quit
objWorkbook.Save
objWorkbook.Close 
objExcel.Quit

msgbox "Done!              ", 0, "Status..."

'Release Memory
Set objExcel = Nothing
Set objWorkbook = Nothing