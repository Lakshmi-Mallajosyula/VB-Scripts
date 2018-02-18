Set myxl = createobject("excel.application")
 
'Make sure that you have created an excel file before exeuting the script. 
'Use the path of excel file in the below code
'Also make sure that your excel file is in Closed state
myxl.Workbooks.Open "D:\Lakshmi sumalatha\Study material\VB Scripts\Sample VB Scripts\Excel\vbtest.xlsx"
 
myxl.Application.Visible = true

'this is the name of  Sheet  in Excel file "qtp.xls"   where data needs to be entered 
set mysheet = myxl.ActiveWorkbook.Worksheets("Sheet1")
 
'Get the max row occupied in the excel file 
Row=mysheet.UsedRange.Rows.Count

'Get the max column occupied in the excel file 
Col=mysheet.UsedRange.columns.count
 
'To read the data from the entire Excel file
For  i= 1 to Row
    For j=1 to Col
        Msgbox  mysheet.cells(i,j).value
    Next
Next
 
'Save the Workbook
myxl.ActiveWorkbook.Save
 
'Close the Workbook
myxl.ActiveWorkbook.Close
 
'Close Excel
myxl.Application.Quit
 
Set mysheet =nothing
Set myxl = nothing