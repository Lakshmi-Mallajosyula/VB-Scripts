Set FSO = Createobject("Scripting.FileSystemObject")

Set File = FSO.OpenTextFile("E:\Lakshmi\Personal\Study material\2. VB Scripting\Sample VB Scripts\FSO\test1.txt", 1, True)

a = inputbox ("Select the options: 1 - Read a character; 2 - Read a line; 3 - read the whole content")

Select case a
	
	case "1"
		Skipchar = Inputbox ("Enter the starting position ")
		chars = inputbox ("Enter the number of characters to display")	
		
		SkipChar1 = skipchar - 1
		
		file.skip(skipchar1)
		msgbox file.read(chars)
	
	Case "2"
		Line = inputbox ("Enter the line to be displayed")
		skip = line - 1
		for i = 1 to skip
			if file.AtEndOfStream <> True Then
				file.skipline
			Else
				Msgbox ("No data on the specified line number")
				wscript.quit
			End IF	
		Next		
		msgbox file.readline
	
	Case "3"
		msgbox file.readall

	Case Else
		Msgbox ("Invalid Input")
			
End select

Set FSO = Nothing
Set File = Nothing