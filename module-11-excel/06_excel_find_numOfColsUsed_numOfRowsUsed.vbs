'------------------------------------------------------------
' We will Explore:
'	- [Excel.Application] -> [Workbook] -> [Sheets] -> [Rows, Columns, Cells]
'	- Count number of rows used
'	- Count number of columns used
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
Dim rowCount, columnCount

strDirectory = "D:\GitBucket\vbscripting-course\02-indev\excel\tmp_data"
excelDocName = "Sample5.xlsx"
sheetName = "Data2"

'Excel Object
Set objExcel = CreateObject("Excel.Application")

'Excel Document Visible Property
objExcel.visible = False

'Workbook Object
Set objWorkbook = objExcel.Workbooks.Open(strDirectory & "\" &  excelDocName) 

'Get the Number of Rows used in the Excel sheet
rowCount = objWorkbook.Sheets(sheetName).UsedRange.Rows.count
msgbox "Number of rows used in sheet " & sheetName & " : " & rowCount

'Get the Number of Columns used in the Excel sheet
columnCount = objWorkbook.Sheets(sheetName).UsedRange.Columns.count
msgbox "Number of columns used in sheet "  & sheetName & " : " & columnCount

'Close & Quit
objWorkbook.Close 
objExcel.Quit

'Release Memory
Set objExcel = Nothing
Set objWorkbook = Nothing