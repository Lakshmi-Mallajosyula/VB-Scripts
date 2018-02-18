Dim fso, file, file_location

file_location = "D:\Lakshmi sumalatha\Study material\VB Scripts\Sample VB Scripts\FSO\test1.txt"

Set fso = CreateObject("Scripting.FileSystemObject")

fso.DeleteFile(file_location)



