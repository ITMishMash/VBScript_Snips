FUNCTION GetDateTime(intOption, chDateSeparator, chTimeSeparator)
'**************************************************************************************
'Declare the type of return using an integer for intOption
'The case statement at the end will build a date string based on the intOption
'Currently only two options created
'SAMPLE: strTimeStamp = GetDateTime(1, FALSE, "?") returns verbose date string, separators are not used
'SAMPLE: strTimeStamp = GetDateTime(2, "?", False) returns date and time stamp, dates are separated by '?'

	Const vbSunday = 1
	Const vbMonday = 2
	Const vbTuesday = 3
	Const vbWednesday = 4
	Const vbThursday = 5
	Const vbFriday = 6
	Const vbSaturday = 7

	Const vbJanuary = 1
	Const vbFebruary = 2
	Const vbMarch = 3
	Const vbApril = 4
	Const vbMay = 5
	Const vbJune = 6
	Const vbJuly = 7
	Const vbAugust = 3
	Const vbSeptember = 4
	Const vbOctober = 5
	Const vbNovember = 6
	Const vbDecember = 7

	Dim dtNow
	Dim strDate, strTime, strDayOfWeek, strMonthOfYear, strTimeStampString
	Dim intDayCode, intMonthCode
	
	dtNow = Now()
	
	If Not(Not(chDateSeparator = FALSE)) Then
		chDateSeparator = "-"
	End If
	
	If Not(Not(chTimeSeparator = FALSE)) Then
		chTimeSeparator = "."
	End If

	strDate = Right("0" & DatePart("m",dtNow), 2) & chDateSeparator _
			& Right("0" & DatePart("d",dtNow), 2) & chDateSeparator _ 
			& DatePart("yyyy",dtNow)

	strTime = Right("0" & DatePart("h",dtNow), 2) & chTimeSeparator _
			& Right("0" & DatePart("h",dtNow), 2) & chTimeSeparator _ 
			& Right("0" & DatePart("s",dtNow), 2)
			
	intDayCode = DatePart("w", dtNow)
	Select Case intDayCode
		Case vbSunday		strDayOfWeek = "Sunday"
		Case vbMonday		strDayOfWeek = "Monday"
		Case vbTuesday		strDayOfWeek = "Tuesday"
		Case vbWednesday	strDayOfWeek = "Wednesday"
		Case vbThursday		strDayOfWeek = "Thursday"
		Case vbFriday		strDayOfWeek = "Friday"
		Case vbSaturday		strDayOfWeek = "Saturday"
	End Select
	
	intMonthCode = DatePart("m", dtNow)
	Select Case intMonthCode
		Case vbJanuary		strMonthOfYear = "January"
		Case vbFebruary		strMonthOfYear = "February"
		Case vbMarch		strMonthOfYear = "March"
		Case vbApril		strMonthOfYear = "April"
		Case vbMay		strMonthOfYear = "May"
		Case vbJune		strMonthOfYear = "June"
		Case vbJuly		strMonthOfYear = "July"
		Case vbAugust		strMonthOfYear = "August"
		Case vbSeptember	strMonthOfYear = "September"
		Case vbOctober		strMonthOfYear = "October"
		Case vbNovember		strMonthOfYear = "November"
		Case vbDecember		strMonthOfYear = "December"
	End Select
	
	Select Case intOption
		Case 1	strTimeStampString = strDayOfWeek & ", " & strMonthOfYear & " " & DatePart("d",dtNow) & ", " & DatePart("yyyy",dtNow)
		Case 2 strTimeStampString = strDate & "_" & strTime
	End Select
	
	GetTime = strTimeStampString
			
'**************************************************************************************
END FUNCTION
