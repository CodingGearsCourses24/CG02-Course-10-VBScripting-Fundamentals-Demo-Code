'------------------------------------------------------------
' We will Explore:
'	- [Excel.Application] -> [Workbook] -> [Sheets] -> [Rows, Columns, Cells]
'	- Create an Microsoft Excel Document
'	- Add new worksheets - MyTemp1 & MyTemp2
'	- Add some data into the new sheets
'	- Count the number of sheets in the workbook
'	- Display the names of the sheets
'------------------------------------------------------------

Option Explicit

Dim objExcel
Dim objWorkbook
Dim objWorkSheet
Dim ObjSheet

Dim sheetName
Dim strDirectory
Dim excelDocName

Dim numOfSheets
Dim a

strDirectory = "D:\GitBucket\vbscripting-course\02-indev\excel\tmp_1"
excelDocName = "MyExcelDocument4.xlsx"

'Excel Object
Set objExcel = CreateObject("Excel.Application")

'Excel Document Visible Property
objExcel.visible = False

'Workbook Object
Set objWorkbook = objExcel.Workbooks.Add 

'Add a new sheet with a name
objWorkbook.WorkSheets.Add.Name = "MyTemp1"
objWorkbook.WorkSheets("MyTemp1").Cells(1,1).value = "Sheet1" 

'Add a new sheet with a name
objWorkbook.WorkSheets.Add.Name = "MyTemp2"
objWorkbook.WorkSheets("MyTemp2").Cells(1,1).value = "Sheet2" 

'Count Sheets
numOfSheets = objWorkbook.Sheets.Count

msgbox "Number of sheets : " & numOfSheets, 0, "Information:"

For  a = 1 To numOfSheets Step 1
  sheetname = objWorkbook.Sheets(a).Name
  msgbox a & "of " & numOfSheets & " sheet name : " & sheetname, 0, "Information:"
Next

'Save
objWorkbook.SaveAs strDirectory & "\" &  excelDocName

'Close, Quit
objWorkbook.Close 
objExcel.Quit

'Release Memory
Set objExcel = Nothing
Set objWorkbook = Nothing