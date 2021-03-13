'------------------------------------------------------------
' Script Overview:
'	- Create a excel document
'	- Add two new sheets - MyTemp1 & MyTemp2
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
excelDocName = "MyExcelDocument4.xlsx"

'Excel Object
Set objExcel = CreateObject("Excel.Application")

'Excel Document Visible Property
objExcel.visible = True

'Workbook Object 
Set objWorkbook = objExcel.Workbooks.Add 

'Add a new sheet with a name & adding some text [Cells(Row,Column)]
objWorkbook.Sheets.Add.Name = "MyTemp1"

'Adding Data on to row 1 (Header)
objWorkbook.Sheets("MyTemp1").Range("A1").value = "Student ID" 
objWorkbook.Sheets("MyTemp1").Range("B1").value = "Name" 
objWorkbook.Sheets("MyTemp1").Range("C1").value = "Grade" 
objWorkbook.Sheets("MyTemp1").Range("D1").value = "Age" 
objWorkbook.Sheets("MyTemp1").Range("E1").value = "Height" 

'Adding Data on to row 2 (Data)
objWorkbook.Sheets("MyTemp1").Range("A2").value = "20180089" 
objWorkbook.Sheets("MyTemp1").Range("B2").value = "John Patri" 
objWorkbook.Sheets("MyTemp1").Range("C2").value = "2nd" 
objWorkbook.Sheets("MyTemp1").Range("D2").value = "9" 
objWorkbook.Sheets("MyTemp1").Range("E2").value = "4'1''" 

'Adding Data on to row 3 (Data)
objWorkbook.Sheets("MyTemp1").Range("A3").value = "20180090" 
objWorkbook.Sheets("MyTemp1").Range("B3").value = "Larry Weel" 
objWorkbook.Sheets("MyTemp1").Range("C3").value = "2nd" 
objWorkbook.Sheets("MyTemp1").Range("D3").value = "9" 
objWorkbook.Sheets("MyTemp1").Range("E3").value = "4'2''" 

'Adding Data on to row 4 (Data)
objWorkbook.Sheets("MyTemp1").Range("A4").value = "20180091" 
objWorkbook.Sheets("MyTemp1").Range("B4").value = "Mary Dat" 
objWorkbook.Sheets("MyTemp1").Range("C4").value = "2nd" 
objWorkbook.Sheets("MyTemp1").Range("D4").value = "9" 
objWorkbook.Sheets("MyTemp1").Range("E4").value = "3'9''" 

'Adding Data on to row 5 (Data)
objWorkbook.Sheets("MyTemp1").Range("A5").value = "20180092" 
objWorkbook.Sheets("MyTemp1").Range("B5").value = "Matt Peterson" 
objWorkbook.Sheets("MyTemp1").Range("C5").value = "3rd" 
objWorkbook.Sheets("MyTemp1").Range("D5").value = "9" 
objWorkbook.Sheets("MyTemp1").Range("E5").value = "4'3''" 

'Formatting (By Range A1 thru E1)
objWorkbook.Sheets("MyTemp1").Range("A1:E1").Font.Name = "Cambria"
objWorkbook.Sheets("MyTemp1").Range("A1:E1").Font.Bold = True
objWorkbook.Sheets("MyTemp1").Range("A1:E1").Font.Italic = True
objWorkbook.Sheets("MyTemp1").Range("A1:E1").Font.Size = 15
objWorkbook.Sheets("MyTemp1").Range("A1:E1").Font.ColorIndex = 5 ' Blue

'Formatting (By Range A2 thru E5)
objWorkbook.Sheets("MyTemp1").Range("A2:E5").Font.Name = "Cambria"
objWorkbook.Sheets("MyTemp1").Range("A2:E5").Font.Bold = False
objWorkbook.Sheets("MyTemp1").Range("A2:E5").Font.Italic = False
objWorkbook.Sheets("MyTemp1").Range("A2:E5").Font.Size = 10
objWorkbook.Sheets("MyTemp1").Range("A2:E5").Font.ColorIndex = 10 ' Green

'Formating Coumn Width
objWorkbook.Sheets("MyTemp1").Columns("A").AutoFit()
objWorkbook.Sheets("MyTemp1").Columns("B").AutoFit()
objWorkbook.Sheets("MyTemp1").Columns("C").AutoFit()
objWorkbook.Sheets("MyTemp1").Columns("D").AutoFit()
objWorkbook.Sheets("MyTemp1").Columns("E").ColumnWidth=30

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