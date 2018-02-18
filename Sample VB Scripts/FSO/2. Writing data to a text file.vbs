Const ForReading = 1, ForWriting = 2, ForAppending = 8

Dim fso, file, file_location

file_location = "E:\Lakshmi\Personal\Study material\2. VB Scripting\Sample VB Scripts\FSO\test1.txt"


Set fso = CreateObject("Scripting.FileSystemObject")
set file = fso.OpenTextFile(file_location, 2, true)

file.Writeline "This is a place to get all your qtp"
file.Writeline "questions and answers solved."

file.write "test test test"
file.write "test2 test2 test2 "

file.writeblanklines(5)

set file = fso.OpenTextFile(file_location, 1, true)

Do While file.AtEndOfStream <> True
	msgbox file.readline
Loop

file.close


set file = nothing
set FSO = nothing
