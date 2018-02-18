Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim fso, file, file_location

file_location = "D:\Lakshmi sumalatha\Study material\VB Scripts\Sample VB Scripts\FSO\test1.txt"

Set fso = CreateObject("Scripting.FileSystemObject")
set file = fso.OpenTextFile(file_location, ForAppending, true)

file.Writeline "This is a new set of data"

file.Write "It is appended to the existing file"

set file = fso.opentextfile(file_location, 1, true)

do while file.AtEndOfStream <> True
	Msgbox file.readall
Loop

set file = nothing
set FSO = nothing