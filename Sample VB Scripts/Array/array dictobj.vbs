Dim a
a = array(2,3,4,5)
Set dobj = createobject("Scripting.Dictionary")
dobj.Add "mykey", a


msgbox dobj("mykey")(1)

count = dobj.count
i = dobj.items
j = dobj.keys

For x = 0 to count-1
	msgbox  i(x) & " :" & j(x)

Next