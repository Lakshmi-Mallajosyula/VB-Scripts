maths = inputbox ("Enter marks in maths")
science = inputbox ("Enter marks in science")
social = inputbox ("Enter marks in social")
english = inputbox ("Enter marks in English")

Total_marks = cdbl(maths) + cdbl(science) + cdbl(social) + cdbl(english)
msgbox "Total marks of the student is " &Total_marks

Avg_Marks = Total_marks/4
msgbox "Average marks of the student is " &Avg_Marks

if Avg_Marks >= 75 then
	Msgbox "The student acquired distinction"
else if Avg_Marks >= 60 and Avg_Marks < 75 then
	Msgbox "The student acquired First class"
else if Avg_Marks >= 50 and Avg_Marks < 60 then
	Msgbox "The student acquired Second class"
else if Avg_Marks >= 40 and Avg_Marks < 50 then
	Msgbox "The student acquired Third class"
else
	Msgbox "The student got Failed!!!"
end if
end if
end if
end if