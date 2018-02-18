dim dictobj

set dictobj = createobject("scripting.dictionary")
dictobj.add "key1","10"
dictobj.add "key2","20"
dictobj.add "key3","30"

msgbox dictobj("key1")

count = dictobj.count
i = dictobj.items
j = dictobj.keys

For x = 0 to count-1
	msgbox  i(x) & " :" & j(x)

Next




'if dictobj.exists("key2") Then 
'	dictobj.remove "key2"
'end if

'msgbox dictobj.count

dictobj.removeall

'msgbox dictobj.count