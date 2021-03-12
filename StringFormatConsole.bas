'#CONSOLE ON
#Define UNICODE
#Include Once "windows.bi"

CONST SOME_FORMAT as string = "test1 %1 test2 %2 test3 %3 test4 %1 test5"
Dim someStringArray(1 to 3) as String = {"TESTA", "TESTB", "TESTC"}

' --------------------------------
' Replace symbols in format string with values in array
'
Private Function StringFormat(formatString as string,  valueArray() as String) as string
    Dim returnValue as String = ""
    Dim workString as string
    Dim symbolString as string
    const SYMBOL as string = "%"
    Dim foundPosition as integer = 0
    Dim leftString as string = ""
    Dim rightString as string = ""
    
    workString=formatString
    for index as integer = LBound(valueArray) to UBound(valueArray)
        symbolString = SYMBOL & Str(index)
        foundPosition = 0
        'For every item in valueArray (using index), 
        ' look for replacement symbol in formatString with same number ('%n' where 'n' is index) and replace the symbol with that array item.
        ' If no index match is found for a given array item, no replacement will be performed with that item.
        ' If a replacement symbol does not have a corresponding array item, it will not be replaced.
        'check 1st time
        foundPosition = instr(workString, symbolString)
        'Consider exhaustive search/replace with each array item for multiple matching symbols.
        Do While foundPosition > 0
            leftString = Left(workString, foundPosition-1)
            rightString=Right(workString, Len(workString)-foundPosition-1)
            workString = leftString & valueArray(index) &rightString
            'check again
            foundPosition = instr(workString, symbolString)
        loop

        'workString = workString & symbolString & valueArray(index)
    next
    
    returnValue = workString
    
    StringFormat = returnValue
    'Exit Function
end function
'
' --------------------------------
'

Print StringFormat(SOME_FORMAT, someStringArray())

Print "Press any key..."
Sleep '3000
