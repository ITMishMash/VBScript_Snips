Function CalculateMaxColumnLength(arrMyArray)
'Return the count of the longest row for each column in an array
'Requires that the array be square and in the format Col, Row
	Dim temp
	Dim cols, rows
	Dim arrLongest()
	cols = UBound(arrMyArray, 1)
	rows = UBound(arrMyArray, 2)
	ReDim arrLongest(cols)
	
	For i = 0 to cols
		temp = 0
		For j = 0 to rows
			If len(arrMyArray(i, j)) > temp Then
				temp = len(arrMyArray(i, j))
			End If
		Next
		arrLongest(i) = temp
	Next
	CalculateMaxColumnLength = arrLongest
End Function

'Sample usage:
'Create an example array.
Dim arrThisArray(2,3)
arrThisArray(0,0) = "Apple" 
arrThisArray(0,1) = "Orange"
arrThisArray(0,2) = "Grapes"           
arrThisArray(0,3) = "pineapple" 
arrThisArray(1,0) = "cucumber"           
arrThisArray(1,1) = "beans"           
arrThisArray(1,2) = "carrThisArrayot"           
arrThisArray(1,3) = "tomato"           
arrThisArray(2,0) = "potato"             
arrThisArray(2,1) = "sandwitch"            
arrThisArray(2,2) = "coffee"             
arrThisArray(2,3) = "nuts"  
'Create an array to hold the results
Dim arrMaxLength
'Call the function
arrMaxLength = CalculateMaxColumnLength(arrThisArray)
'Loop through the results and compare them to what was supplied
dim msg
msg = Ubound(arrMaxLength) + 1 & " total columns" & vbCrlf
For i = 0 to Ubound(arrMaxLength)
	msg = msg + "Column " & i + 1 & " is a max of  " & arrMaxLength(i) & " characters long." & vbCrlf
Next
msgbox msg
