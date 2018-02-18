set myxls = createobject("excel.application")

myxls.workbooks.add

myxls.activeworkbook.worksheets.delete
myxls.activeworkbook.saveas ("D:\Lakshmi sumalatha\Study material\VB Scripts\Sample VB Scripts\Excel\addsheetexcel.xlsx")

myxls.activeworkbook.close

myxls.application.quit

set myxls = nothing