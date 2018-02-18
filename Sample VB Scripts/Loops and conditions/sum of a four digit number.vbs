a = inputbox ("Enter a four digit number")
sum = 0

if a = empty or len(a) <> 4 then
	msgbox "Please enter a valid number"

else
	for i = 1 to 4
		value = mid(a, i, 1)
		sum = sum + value
	next
	msgbox "The sum of the four digits in the number: " &a &" is " &sum
	msgbox "The reverse of the four digits in the number: " &a &" is " &strreverse(a)

end if



