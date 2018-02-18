a = inputbox ("Enter a ten digit number")

if a = empty or len(a) <> 10 then
	msgbox "Please enter a valid number"
	wscript.quit
end if

for i = 1 to 10
	value = mid(a, i, 1)
	result = isnumeric(value)
	if result = false then
		msgbox "The input given is not numeric"
		WScript.quit 
	end if	
next

first_num = left(a, 1)

if first_num = 9 or first_num = 8 then
	msgbox "Entered number is a valid mobile number"

else
	msgbox "Entered number is not a valid mobile number"
end if
