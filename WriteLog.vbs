'WriteLogFile will write a message to a log file given the path of the file and the message
'WriteLogDB will write a message to a DB given the connection string and the message

SUB WriteLogFile(strLogFilePath, strLogMessage)
'**************************************************************************************
'SAMPLE USAGE: WriteLogFile "C:my_logfile.log" "Log message"
	Const ForReading = 1
	Const ForWriting = 2
	Const ForAppending = 8

	Dim objFS, objLogFile

	Set objFS = CreateObject("Scripting.FileSystemObject")
	Set objLogFile = objFS.OpenTextFile(LOG_FILE, ForAppending, True)

	objLogFile.WriteLine(Now() & ": " & strLogMessage)
	Set objLogFile = Nothing
	Set objFS = Nothing
'**************************************************************************************
END SUB

SUB WriteLogDB(strConnectionString, strLogTable, arrLogMessage, arrTableHeaders)
'**************************************************************************************
'arrTableHeaders is optional and can be bypassed by entering FALSE in its place

'SAMPLE USAGE: 
	'Const strServerName="myServerName"
	'Const strDBuid="myUserName"
	'Const strDBpw="myPW"
	'Const strDBName="myDBName"
	'Dim strMyConnectionString, strMyLogTable
	'Dim arrMyTableHeaders, arrMyLogMessage
	'arrMyTableHeaders = Array("Field1","Field2","Field3")
	'ReDim arrMyLogMessage(UBound(arrTableHeaders))
	'
	'arrLogMessage(0) = "Field1 Message Values"
	'arrLogMessage(1) = "Field2 Message Values"
	'arrLogMessage(2) = "Field3 Message Values"
	'' Alternatively arrLogMessage can be populated by using arrLogMessage = Array("Field1 Message Values","Field2 Message Values","Field3 Message Values")
	'strMyConnectionString  "driver={SQL SERVER};server=" & strServerName & ";uid=" & strDBuid & ";pwd=" & strDBpw & ";database=" & strDBName
	'strMyLogTable = “myLogTableName”
	'WriteLogDB strMyConnectionString strMyLogTable arrMyLogMessage arrMyTableHeaders

	Dim objConnection
	Dim strLogValues, strHeaderValues, strQuery
	Dim boolHeaders
	Dim intColumns, i
	
	'Check to see if the table headers were passed in and declare a flag
	If Not(arrTableHeaders = FALSE) Then
		boolHeaders = TRUE
	Else
		boolHeaders = FALSE
	End If
	
	'Store the number of columns as an int
	intColumns = UBound(arrTableHeaders)
	'Start the Values insert statement
	strLogValues = "("
	
	For i = 0 to intColumns
		strLogValues = strLogValues & "'" &arrLogMessage(i) & "'"
		If i < intColumns Then
			strLogValues = strLogValues & ','
		Else
			strLogValues = strLogValues & ');'
		End If
	Next
	
	If boolHeaders Then
		'strQuery Sample: INSERT INTO myLogTableName (Field1,Field2,Field3) VALUES ('myField1Value','myField2Value','myField3Value');
		'strQuery = "INSERT INTO " & strLogTable & " (Field1,Field2,Field3)  VALUES ('" & Now() & "','" & "('" & strLogMessage & & "','" "')"
		Dim strHeaderValues
		strHeaderValues = "("
		For i = 0 to intColumns
			strHeaderValues = strHeaderValues & arrTableHeaders(i)
			If i < intColumns Then
				strHeaderValues = strHeaderValues & ','
			Else
				strHeaderValues = strHeaderValues & ')'
			End If
		Next
		strQuery = "INSERT INTO " & strLogTable & " " & strHeaderValues & " VALUES " & strLogValues
	Else
		strQuery = "INSERT INTO " & strLogTable & " VALUES " & strLogValues
	End If
	
	On error resume next
	Set objConnection = CreateObject("ADODB.Connection") 
		objConnection.Open(strConnectionString) 
		objConnection.Execute(strQuery) 
		objConnection.Close
	Set objConnection = Nothing
	Erase arrLogMessage
	On error goto 0
'**************************************************************************************
END SUB
