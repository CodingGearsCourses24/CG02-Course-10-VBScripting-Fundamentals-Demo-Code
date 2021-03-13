'---------------------------------------------------------------------
' We will Explore:
'	- [Excel.Application] -> [Workbook] -> [Sheets] -> [Rows, Columns, Cells]
'	- Create an Microsoft Excel Document
'	- Add new worksheets - MyTemp1 thru MyTemp4
'   - By default, excel adds new sheets to the left of the active sheet.
'   - Here..... Sheets are added at the end of the existing sheets.
'			Code Snippet: objExcel.Sheets.Add(objWorkbook.Sheets(objWorkbook.Sheets.Count)).Name = "MyTemp2"
'	- Add some data into the new sheets
'	- Delet the default sheet "sheet1"
'	- Save the document as ""MyExcelDocument1.xlsx"
'---------------------------------------------------------------------

Option Explicit

Dim objExcel
Dim objWorkbook
Dim objWorkSheet, ws
Dim ObjSheet

Dim sheetName
Dim strDirectory
Dim excelDocName

strDirectory = "D:\New folder\tmp_1"
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
objExcel.Sheets.Add(objWorkbook.Sheets(objWorkbook.Sheets.Count)).Name = "MyTemp1"
objExcel.Sheets("MyTemp1").Cells(1,1).value = "Using Cells" 
objExcel.Sheets("MyTemp1").Cells(5,1).value = "Using Cells 51" 
objExcel.Sheets("MyTemp1").Range("A2").value = "Using Range" 

'Add another new sheet with a desired name & adding some text
objExcel.Sheets.Add(objWorkbook.Sheets(objWorkbook.Sheets.Count)).Name = "MyTemp2"
objExcel.Sheets("MyTemp2").Cells(1,1).value = "Using Cells" 
objExcel.Sheets("MyTemp2").Range("A2").value = "Using Range" 

'Add another new sheet with a desired name & adding some text
objExcel.Sheets.Add(objWorkbook.Sheets(objWorkbook.Sheets.Count)).Name = "MyTemp3"
objExcel.Sheets("MyTemp3").Cells(1,1).value = "Using Cells" 
objExcel.Sheets("MyTemp3").Range("A2").value = "Using Range" 

'Add another new sheet with a desired name & adding some text
objExcel.Sheets.Add(objWorkbook.Sheets(objWorkbook.Sheets.Count)).Name = "MyTemp4"
objExcel.Sheets("MyTemp4").Cells(1,1).value = "Using Cells" 
objExcel.Sheets("MyTemp4").Range("A2").value = "Using Range" 

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