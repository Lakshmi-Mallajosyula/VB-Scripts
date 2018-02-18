
set myxls = createobject("excel.application")

myxls.application.visible = True

myxls.workbooks.open("E:\Lakshmi\Personal\Study material\2. VB Scripting\Sample VB Scripts\Excel\vbtest.xlsx")

set mysheet = myxls.activeworkbook.worksheets("Sheet1")

mysheet.cells(2,1).value = "Testdata"
mysheet.cells(2,2).value = "Test2"

row = mysheet.usedrange.rows.count
columns = mysheet.usedrange.columns.count

msgbox "Number of rows: " & row &" and Number of columns: " & columns

myxls.activeworkbook.save

myxls.activeworkbook.close
myxls.application.quit

set myxls = nothing

set mysheet = nothing