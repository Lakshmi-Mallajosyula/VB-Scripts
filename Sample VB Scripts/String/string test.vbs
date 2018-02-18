Var1 = "This is a sample test"


for i = len(Var1) to 1 step -1
	strchar = mid(Var1, i, 1)
	result = result &strchar
next
msgbox result


msgbox strreverse(var1)