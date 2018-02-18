Dim myArray(5)   'Declare Array with 5 elements
Set myArray(0) = createobject("scripting.dictionary")         'Make first element in array as a dictionary object
myArray(0).Add"mykey","myvalue"       'Once we have a dictionary object, We can use its methods like add, remove, removeall etc
print myArray(0)("mykey")          'display item value of mykey in dictionary myArray(0)
myArray(0).removeall 
