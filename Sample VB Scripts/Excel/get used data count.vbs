set myxls = createobject("excel.application")

myxls.workbooks.open "D:\Lakshmi sumalatha\Study material\VB Scripts\Sample VB Scripts\Excel\vbtest.xlsx"

myxls.application.visible = True

set mysheet = myxls.ActiveWorkBook.worksheets("Sheet1")


Row = mysheet.usedrange.rows.count

msgbox Row

Column = mysheet.usedrange.columns.count

msgbox Column

myxls.activeworkbook.save

myxls.activeworkbook.close

myxls.application.quit

set myxls = nothing
set mysheet = nothing