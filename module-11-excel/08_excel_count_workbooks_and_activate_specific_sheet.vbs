'------------------------------------------------------------
' We will Explore:
'	- Count open workbooks
'	- Activate a desired worksheet in a workbook
'------------------------------------------------------------

Option Explicit

Dim objExcel
Dim objWorkbook
Dim objWorkSheet
Dim ObjSheet

Dim sheetName
Dim strDirectory
Dim objWorkbook1
Dim objWorkbook2
Dim excelDoc1
Dim excelDoc2
Dim workbookCount

strDirectory = "D:\GitBucket\vbscripting-course\02-indev\excel\tmp_data"
excelDoc1 = "Sample1.xlsx"
excelDoc2 = "Sample2.xlsx"

'Excel Object
Set objExcel = CreateObject("Excel.Application")

'Excel Document Visible Property
objExcel.visible = True

'Getting a workbook object (using specific file)
Set objWorkbook1 = objExcel.Workbooks.Open(strDirectory & "\" &  excelDoc1) 
Set objWorkbook2 = objExcel.Workbooks.Open(strDirectory & "\" &  excelDoc2) 

workbookCount = objExcel.Workbooks.count

MsgBox "Total number of open workbooks : " & workbookCount, 0, "Workbooks"

objExcel.Workbooks("Sample1.xlsx").Worksheets("Sheet1").Activate

'Replace line 42 with "objWorkbook1.Worksheets("Sheet1").Activate". Will it work?

MsgBox "Do NOT click Okay. Verify that the Sheet1 in Sample1 workbook is active on your screen"

'Close & Quit
objWorkbook1.Close
objWorkbook2.Close 
objExcel.Quit