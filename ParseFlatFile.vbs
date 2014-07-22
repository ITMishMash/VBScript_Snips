SUB ParseRecords()
'**************************************************************************************
'Read a flat file, parse it into an array, and release the file
'TODO: Refactor to accept params

	Dim strCurrentLine()
	Dim arrCurrentRecord
	Dim strLine
	Dim intCurrentRow, intCurrentCol

	intCurrentRow = 0
	intCurrentCol = 0

	'Loop through the output.txt file and assign the values to an array
	Do Until webservicefile.atendofstream
		'Resize the array for the next line
		Redim Preserve strCurrentLine(intCurrentRow)
		'Have to escape out the apostrophes
		strCurrentLine(intCurrentRow) = replace(webservicefile.readline,"'","''")
		intCurrentRow = intCurrentRow + 1
	        'Store the total number of rows
		intNumRows = intCurrentRow - 1
	Loop
	
	'Create the array to hold all of the values
	Redim arrRecords (intNumRows, (intNumCols-1))
	'wscript.echo "Created array of "  & intNumRows + 1 & " Rows and " & intNumCols & " Columns"
	'Store the values in the array
	For intCurrentRow = 0 to intNumRows
		arrCurrentRecord = Split(strCurrentLine(intCurrentRow),"|")
		For intCurrentCol = 0 to (intNumCols - 1)
			'wscript.echo "Parsing " & arrCurrentRecord(intCurrentCol) & " to " & intCurrentRow & "," & intCurrentCol
			arrRecords(intCurrentRow,intCurrentCol) = arrCurrentRecord(intCurrentCol)
			'intCurrentCol = intCurrentCol + 1
		Next
	Next
	'Release the textfile array from memory
	Erase strCurrentLine
'**************************************************************************************
END SUB
