set myxls = createobject("excel.application")

myxls.workbooks.add

set mysheet = myxls.activeworkbook.worksheets("Sheet1")

mysheet.cells(1,1).value = "Test"
mysheet.cells(1,2).value = "Test2"
mysheet.cells(2,1).value = "Test3"
mysheet.cells(2,2).value = "Test4"

myxls.activeworkbook.saveas ("D:\Lakshmi sumalatha\Study material\VB Scripts\Sample VB Scripts\Excel\new updated.xlsx")

myxls.activeworkbook.close

myxls.application.quit

set myxls = nothing
set mysheet = nothing