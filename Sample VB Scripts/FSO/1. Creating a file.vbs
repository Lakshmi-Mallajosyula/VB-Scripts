Dim fso, file, file_location

file_location = "E:\Lakshmi\Personal\Study material\2. VB Scripting\Sample VB Scripts\FSO\test1.txt"

Set fso = CreateObject("Scripting.FileSystemObject")

set file = fso.createTextFile(file_location)

file.writeline "This is a test"




