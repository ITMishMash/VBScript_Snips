Function ParseFlatFile(strFileName, boolDropLastColumn)		'Read a flat file, parse it into an array, and release the file
	Dim objFSO, objFile
	Dim strCurrentLine
	Dim arrLines(), arrRecords
	Dim arrCurrentRecord
	Dim strLine
	Dim intCurrentRow, intCurrentCol
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(strFileName, ForReading)
	intCurrentRow = 0
	intCurrentCol = 0

	Do Until objFile.atendofstream								'Loop through the file and assign the values to an array	
		'strCurrentLine(intCurrentRow) = replace(objFile.readline,"'","''") 'Use this to escape out the apostrophes
		strCurrentLine = objFile.ReadLine
		If Not(strCurrentLine = "") Then
			Redim Preserve arrLines(intCurrentRow)		'Resize the array for the next line
			arrLines(intCurrentRow) = strCurrentLine
			intCurrentRow = intCurrentRow + 1
		End If
	Loop
	objFile.Close
	Set objFSO = Nothing
	
	intNumRows = Ubound(arrLines)							'Hold the total number of rows
	intNumCols = UBound(Split(arrLines(0),"|"))			'Hold the total number of columns
	If boolDropLastColumn Then								'Check to see if we should store the last column
		intNumCols = intNumCols - 1
	End If

	Redim arrRecords (intNumCols, intNumRows)								'Create the array to hold all of the values
	'wscript.echo "Created array of "  & intNumRows + 1 & " Rows and " & intNumCols & " Columns"
	For intCurrentRow = 0 to intNumRows											'Store the values in the array
		arrCurrentRecord = Split(arrLines(intCurrentRow),"|")
		For intCurrentCol = 0 to (intNumCols)
			'wscript.echo "Parsing " & arrCurrentRecord(intCurrentCol) & " to " & intCurrentRow & "," & intCurrentCol
			arrRecords(intCurrentCol, intCurrentRow) = arrCurrentRecord(intCurrentCol)
		Next
	Next
	Erase arrLines																	'Release the textfile array from memory
	ParseFlatFile = arrRecords															'Return the record array
End Function
