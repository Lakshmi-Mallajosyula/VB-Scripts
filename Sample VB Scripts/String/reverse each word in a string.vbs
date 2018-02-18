a = "this is a test vb script. vb script is easy to learn. all you need for vb script is a right attitude"

b = split(a)

for each x in b
	
	y = y &" " &strreverse(x)
		
next

msgbox "reverse of individual words: "
msgbox y


for i = len(a) to 1 step -1
	strchar = mid(a, i, 1)
	result = result &strchar
next
msgbox "reverse of sentence: "
msgbox result