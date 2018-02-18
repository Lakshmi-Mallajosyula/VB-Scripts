
set myxls = createobject("excel.application")

myxls.application.visible = true

set workbook1 = myxls.workbooks.open("D:\Lakshmi sumalatha\Study material\VB Scripts\Sample VB Scripts\Excel\Test1.xlsx")

set sheet1 = workbook1.worksheets("sheet1")

set workbook2 = myxls.workbooks.open("D:\Lakshmi sumalatha\Study material\VB Scripts\Sample VB Scripts\Excel\Test2.xlsx")

set sheet2 = workbook2.worksheets("sheet1")


for each cell in sheet1.usedrange

	if cell.value <> sheet2.range(cell.address).value then
		cell.interior.colorindex = 3
		mismatch = 1
		workbook1.save
	end if

next

for each cell in sheet2.usedrange

	if cell.value <> sheet1.range(cell.address).value then
		cell.interior.colorindex = 4
		mismatch = 1
		workbook2.save
	end if

next

if mismatch = 0 then
	msgbox "no mismatch"
end if

Workbook1.close
Workbook2.close
 
myxls.application.Quit

set myxls=nothing