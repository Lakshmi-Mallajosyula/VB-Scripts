Dim fso, file, file_location, b(10)

file_location = "D:\Lakshmi sumalatha\Study material\VB Scripts\Sample VB Scripts\FSO\test1.txt"

Set fso = CreateObject("Scripting.FileSystemObject")

set file = fso.opentextfile(file_location, 1, true)

a = inputbox("enter the line number to view: ")

line = a - 1

for i = 1 to line
	if file.AtEndOfStream <>true then
		file.skipline
	else
		msgbox "out of lines. No data to display"
		wscript.quit
	end if
next

msgbox "The line number selected to view is: " & file.line
msgbox file.readline


Set FSO = Nothing
Set File = Nothing




set fso = nothing
set file = nothing




