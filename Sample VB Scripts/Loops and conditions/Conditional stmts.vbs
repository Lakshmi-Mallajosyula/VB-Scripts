dim currmonth, currdate

currmonth = month(now)

msgbox "This is " &currmonth &"th month of the year"

select case currmonth

case 1
msgbox "this is january"

case 2
msgbx " this is feb"

case 3 
msgbox "this is mrch"

case 4
msgbox " this is april"

case 5 
msgbox "this is may"

case 6
msgbox "this is june"

case 7
msgbox " this is july"

case 8 
msgbox "this is aug"

case 9
msgbox " this is sep"

case 10
msgbox "this is oct"

case 11
msgbox " this is nov"

case 12
msgbox " this is dec"

case else
msgbox "invalid month"
end select