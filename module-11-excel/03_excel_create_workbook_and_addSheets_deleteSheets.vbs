'---------------------------------------------------------------------
' We will Explore:
'	- [Excel.Application] -> [Workbook] -> [Sheets] -> [Rows, Columns, Cells]
'	- Create an Microsoft Excel Document
'	- Add new worksheets - MyTemp1 & MyTemp2
'	- Add some data into the new sheets
'	- Delet the default sheet "sheet1"
'	- Save the document as ""MyExcelDocument1.xlsx"
'---------------------------------------------------------------------

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

'Excel Document Visible Property
objExcel.visible = False

'Workbook Object 
'  Add method creats a new workbook & 
'  Returns Workbook object
Set objWorkbook = objExcel.Workbooks.Add 

'Add a new sheet with a desired name & adding some text [Cells(Row,Column)]
objExcel.Sheets.Add.Name = "MyTemp1"
objExcel.Sheets("MyTemp1").Cells(1,1).value = "Using Cells" 
objExcel.Sheets("MyTemp1").Cells(5,1).value = "Using Cells 51" 
objExcel.Sheets("MyTemp1").Range("A2").value = "Using Range" 


'Add another new sheet with a desired name & adding some text
objExcel.Sheets.Add.Name = "MyTemp2"
objExcel.Sheets("MyTemp2").Cells(1,1).value = "Using Cells" 
objExcel.Sheets("MyTemp2").Range("A2").value = "Using Range" 

'Delete the default sheet1 
objExcel.Sheets("sheet1").Delete

'Save
objWorkbook.SaveAs strDirectory & "\" &  excelDocName

msgbox "Done!              " , 0, "Status..."

'Close & Quit
objWorkbook.Close 
objExcel.Quit

'Release Memory
Set objExcel = Nothing
Set objWorkbook = Nothing