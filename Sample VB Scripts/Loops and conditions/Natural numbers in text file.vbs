location = "D:\Suma\Study materials\VB Scripts\test1.txt"

n = inputbox("enter the number: ")
a = 1

for i = 2 to n 
	a = a &" , " &i
next

set fso = createobject("scripting.filesystemobject")
set file = fso.opentextfile(location, 2, true)
file.write(a)
file.close