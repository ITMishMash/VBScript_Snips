SUB InsertRecords()
'Parse an array into SQL Insert statements to create multiple inserts instead of individual
'intNumRows declares the number of rows inserted in a single Insert statement
'Optimal Insert performance depends on the number of rows being inserted simultaneously, the number of columns, and the data types 
'Article discussing performance impact of multiple vs. single inserts:
'https://www.simple-talk.com/sql/performance/comparing-multiple-rows-insert-vs-single-row-insert-with-three-data-load-methods/
'
'TODO: Refactor to accept parameters: strInsertStatement, intNumRows, arrFieldHeaders
'TODO: Create timers to log the performance of Insert statements and the process as a whole
'TODO: Dynamically update intNumRecords based on historical performance data to automtically optimize Inserts

	Dim strInsertStatement 
	Dim intCurrentRow
	Dim i
	Dim intUpperRec
	'Create and Execute SQL Insert statements of intRecordsPerInsert rows
	Do While intCurrentRow <= intNumRows
		strInsertStatement = "INSERT INTO " & table & " (Field1,Field2,Field3) VALUES " 'TODO: Update the sub to accept an array for field names. Are they necessary?
		intUpperRec = (intCurrentRow + (intRecordsPerInsert-1))
		If intUpperRec > intNumRows Then
			intUpperRec = intNumRows
		End If
		For i = intCurrentRow to intUpperRec
			If (intCurrentRow = intNumRows) Or (intCurrentRow = intUpperRec) Then
				'wscript.echo "Current Row = " & intCurrentRow & " intUpperRec = " & intUpperRec & " i = " & i
				strInsertStatement = strInsertStatement & "('" & arrRecords(intCurrentRow,0) & "','" & _
					arrRecords(intCurrentRow, 1) & "','" & _
					arrRecords(intCurrentRow, 2) & "');"
			ElseIf intCurrentRow < intNumRows Then
				'wscript.echo "Current Row = " & intCurrentRow & " intUpperRec = " & intUpperRec & " i = " & i
				strInsertStatement = strInsertStatement & "('" & arrRecords(intCurrentRow,0) & "','" & _
					arrRecords(intCurrentRow, 1) & "','" & _
					arrRecords(intCurrentRow, 2) & "'), "
			End If
			intCurrentRow = intCurrentRow + 1
		Next
		'wscript.echo "Executing Insert:  " & strInsertStatement
		Conn.Execute strInsertStatement
	Loop
	'wscript.echo "Total rows affected:  " & intNumRows + 1
End Sub

'********************************************************************************************************************************************
