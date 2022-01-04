'---------------------------------------------------------------------
' We will Explore:
'	- [Excel.Application] -> [Workbook] -> [Sheets] -> [Rows, Columns, Cells]
'	- Reading Excel Document (Specific Sheet & cell)
'	- Using WorkSheet object
'	- To access a specific sheet, use (a) sheet name or (b) index number
'   - 		1 for 1st sheet, 2 for 2nd sheet, etc.
'	- WorkSheets Property returns a collection of worksheets
'---------------------------------------------------------------------

Option Explicit

Dim objExcel
Dim objWorkbook
Dim objWorkSheet '<---
Dim ObjSheet
Dim sheetName
Dim name, age, email
Dim strDirectory
Dim excelDocName

strDirectory = "D:\GitBucket\vbscripting-course\02-indev\excel\tmp_data"
excelDocName = "Sample1.xlsx"
sheetName = "Data1"

'Create Excel Object
Set objExcel = CreateObject("Excel.Application")

'Excel Document Visible Property
objExcel.visible = False

'Getting an workbook object (using specific file)
Set objWorkbook = objExcel.Workbooks.Open(strDirectory & "\" &  excelDocName) 

'Sheets returns a sheet object
'We are binding to a specific sheet based on the name
Set objWorkSheet = objWorkbook.WorkSheets(sheetName)

'Read data from cells
name = objWorkSheet.Range("A2").Value
msgbox "Name: " &  name, 0, "Reading Excel Document..."

age = objWorkSheet.Range("B2").Value
msgbox "Age: " &  age, 0, "Reading Excel Document..."

email = objWorkSheet.Range("C2").Value
msgbox "Email: " &  email, 0, "Reading Excel Document..."

'Close & Quit
objWorkbook.Close 
objExcel.Quit

'Release Memory
Set objExcel = Nothing
Set objWorkbook = Nothing
Set objWorkSheet =  Nothing