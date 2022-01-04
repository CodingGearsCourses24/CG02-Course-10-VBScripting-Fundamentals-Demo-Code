'---------------------------------------------------------------------
' We will Explore:
'	- [Excel.Application] -> [Workbook] -> [Sheets] -> [Rows, Columns, Cells]
'	- Reading Excel Document (Specific Sheet & cell)
'---------------------------------------------------------------------

Option Explicit

Dim objExcel
Dim objWorkbook
Dim objWorkSheet
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

'Read data from cells (using Sheets property of WorkBook object)
name = objWorkbook.Sheets(sheetName).Range("A2").Value
msgbox "Name: " &  name, 0, "Reading Excel Document..."

age = objWorkbook.Sheets(sheetName).Range("B2").Value
msgbox "Age: " &  age, 0, "Reading Excel Document..."

email = objWorkbook.Sheets(sheetName).Range("C2").Value
msgbox "Email: " &  email, 0, "Reading Excel Document..."

'Close & Quit
objWorkbook.Close 
objExcel.Quit

'Release Memory
Set objExcel = Nothing
Set objWorkbook = Nothing