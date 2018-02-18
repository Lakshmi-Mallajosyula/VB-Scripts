a = inputbox ("Enter a ten digit number")

if a = empty or len(a) <> 10 then
	msgbox "Please enter a valid number"

else
	for i = 1 to 10
		value = mid(a, i, 1)
		result = isnumeric(value)
		if result = false then
			msgbox "The input given is not numeric"
			WScript.quit 
		end if	
	next
end if



