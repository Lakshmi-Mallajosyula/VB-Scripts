

set myxls = createobject("excel.application")

myxls.application.visible = "True"

myxls.workbooks.open "D:\Lakshmi sumalatha\Study material\VB Scripts\Sample VB Scripts\Excel\Test1.xlsx"

set mysheet = myxls.activeworkbook.sheets("sheet1")

mysheet.cell(1,1).value = '1'

myxls.activeworkbook.save

myxls.activeworkbook.close

myxls.application.quit

set myxls = nothing

