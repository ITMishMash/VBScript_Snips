Function CountDimensionsInArray(myArray)
'The purpose of this function is to return an array with dimensions
'of the array that is passed to it in. The length of the returned array
'will be equal to the number of dimensions in the supplied array
'The values in the array will represent the number of records in the 
'supplied array, respectively
	On Error Resume Next							'The only way to know that the end of the dimensions in the supplied array have been reached is to check for an error status
	Dim arrBounds()									'Declare an array to hold the limits of the supplied array
	Dim i, temp											'Declare an iterator and a temp value
	i = 0
	While Err.Number = 0							'Loop until an error is returned
		temp = Ubound(myArray,i+1)			'Assign the upper bound of the current dimension to the temp variable
		If Err.Number = 0 Then						'If an error did not occur in the previous step, 
			Redim Preserve arrBounds(i)		'redim the results array 
			arrBounds(i) = temp						'and assign the value
		End If
		i = i + 1												'Increment the iterator
	Wend
	Err.Clear												'Clear any error codes
	CountDimensionsInArray = arrBounds	'Return the results
End Function

'Sample usage:
'Create an example array of the following dimensions. These numbers are not stored values, but rather the length of each dimension
Dim arrThisArray(20,10,2,2,2,1,1,8)
'Create an array to hold the results
Dim arrDimensions
'Call the function
arrDimensions = CountDimensionsInArray(arrThisArray)
'Loop through the results and compare them to what was supplied
dim msg
msg = Ubound(arrDimensions) + 1 & " total dimensions" & vbCrlf
For i = 0 to Ubound(arrDimensions)
	msg = msg + "Dimension " & i + 1 & " has " & arrDimensions(i) & " elements" & vbCrlf
Next
msgbox msg
