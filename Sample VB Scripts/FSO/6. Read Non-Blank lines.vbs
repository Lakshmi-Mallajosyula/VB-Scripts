File_Location = "D:\Lakshmi sumalatha\Study material\VB Scripts\Sample VB Scripts\FSO\test1.txt"

Set FSO = createobject("Scripting.FileSystemObject")

Set File = FSO.OpenTextFile(File_Location, 1, True)

'To read non-blank line

Do while File.AtEndofline <> true
	
	msgbox file.readline

Loop


Set FSO = Nothing
Set File = Nothing