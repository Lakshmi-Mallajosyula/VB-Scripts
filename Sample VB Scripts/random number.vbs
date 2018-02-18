dim min, max, i

min = 10
max = 20

for i = 1 to 4

	randomize
	randomnumber = (int((max-min+1)*rnd)+min)
	msgbox "Random number generated between " &min & " and " &max& " is " &randomnumber 

next