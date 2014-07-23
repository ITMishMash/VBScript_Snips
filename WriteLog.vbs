'WriteLogFile will write a message to a log file given the path of the file and the message
'WriteLogDB will write a message to a DB given the connection string and the message

SUB WriteLogFile(strLogFilePath, strLogMessage)
'**************************************************************************************
'SAMPLE USAGE: WriteLogFile "C:my_logfile.log" "Log message"
	Const LOG_FILE = "D:\OnBase\BPI\scripts\Log\BPI_OB_HTML_Export_logfile.log"
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
	Dim strValues, strHeaders, strQuery
	Dim boolHeaders
	Dim intColumns, i
	
	If Not(arrTableHeaders = FALSE) Then
		boolHeaders = TRUE
	Else
		boolHeaders = FALSE
	End If
	
	intColumns = UBound(arrTableHeaders)
	strValues = "("
	
	For i = 0 to intColumns
		strValues = strValues & "'" &arrLogMessage(i) & "'"
		If i < intColumns Then
			strValues = strValues & ','
		Else
			strValues = strValues & ')'
		End If
	Next
	
	If boolHeaders Then
		'strQuery Sample: INSERT INTO myLogTableName (Field1,Field2,Field3) VALUES ('myField1Value','myField2Value','myField3Value');
		'strQuery = "INSERT INTO " & strLogTable & " (Field1,Field2,Field3)  VALUES ('" & Now() & "','" & "('" & strLogMessage & & "','" "')"
		Dim strHeaderValues
		strHeaderValues = "("
		For i = 0 to intColumns
			strHeaderValues = strHeaderValues & arrMyTableHeaders(i)
			If i < intColumns Then
				strHeaderValues = strHeaderValues & ','
			Else
				strHeaderValues = strHeaderValues & ')'
			End If
		Next
		strQuery = "INSERT INTO " & strLogTable & " (Field1,Field2,Field3)  VALUES "
	Else
		strQuery = "INSERT INTO " & strLogTable & " VALUES "
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
