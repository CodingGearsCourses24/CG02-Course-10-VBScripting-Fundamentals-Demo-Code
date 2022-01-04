'---------------------------------------------------------------------
' We will Explore:
'	- [Excel.Application] -> [Workbook] -> [Sheets] -> [Rows, Columns, Cells]
'	- Reading Excel Document (Specific Sheet & cell)
'	- Using **sheet object**
'   - Sheets retun a collection of sheets in a workbook
'	- To access a specific sheet, use (a) sheet name or (b) index number
'   - 		1 for 1st sheet, 2 for 2nd sheet, etc.
' 	- Sheets Property returns a collection of sheets
' 	- Sheets collection includes all the sheets (chart sheets and worksheets)
'---------------------------------------------------------------------

Option Explicit

Dim objExcel
Dim objWorkbook
Dim objWorkSheet
Dim ObjSheet ' <-----
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

'We are binding to a specific sheet based on the name
Set objSheet = objWorkbook.Sheets(sheetName)

'Read data from cells
name = objSheet.Range("A2").Value
msgbox "Name: " &  name, 0, "Reading Excel Document..."

age = objSheet.Range("B2").Value
msgbox "Age: " &  age, 0, "Reading Excel Document..."

email = objSheet.Range("C2").Value
msgbox "Email: " &  email, 0, "Reading Excel Document..."

'Close & Quit
objWorkbook.Close 
objExcel.Quit

'Release Memory
Set objSheet = Nothing
Set objExcel = Nothing
Set objWorkbook = Nothing