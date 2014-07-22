SUB SendEmail(strTo, strFrom, strSubject, strBody, strMailServer, strFileAttachment)
'**************************************************************************************
'If no file attachment is required, enter FALSE for the strFileAttachmentField when calling the sub
'Separate multiple email addresses using semicolons ;
'strMailServer can be set to localhost or any viable email server

	Const cdoSendUsingPickup = 1
	Const cdoSendUsingPort = 2
	Const cdoSendUsingExchange = 3
	
	Dim objMessage
	Set objMessage= CreateObject("CDO.Message")            
	
		objMessage.To = strTo
		objMessage.From = strFrom
		objMessage.Subject = strSubject
		objMessage.TextBody = strBody
		If Not(strFileAttachment = FALSE) Then
			objMessage.Addattachment strFileAttachment 
		End If
		objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort
		objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strMailServer
		objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		objMessage.Configuration.Fields.Update
		objMessage.Send

	Set objMessage = Nothing
'**************************************************************************************
END SUB
