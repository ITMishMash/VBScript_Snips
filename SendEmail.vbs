SUB SendEmail(strTo, strFrom, strSubject, strBody, strMailServer, strFileAttachment)
'**************************************************************************************
'If no file attachment is required, enter FALSE for the strFileAttachmentField when calling the sub
'Separate multiple email addresses using semicolons ;
'strMailServer can be set to localhost or any viable email server
'TODO: Update additional settings
'Refer to: http://social.msdn.microsoft.com/Forums/en-US/042bd3ab-4425-4ae9-9fdc-1d94fab4cfd3/moving-from-cdo-to-systemnetmail-cdoanonymous?forum=netfxnetcom

	Const cdoSendUsingPickup = 1	'Send message using the local SMTP service pickup directory.
	Const cdoSendUsingPort = 2	'Send the message using the network (SMTP over the network).
	Const cdoSendUsingExchange = 3
	
	'UNUSED CONST
	Const cdoAnonymous = 0 		'Do not authenticate
	Const cdoBasic = 1 		'basic (clear-text) authentication
	Const cdoNTLM = 2 		'NTLM
		
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
		
		'Name or IP of Remote SMTP Server
		objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strMailServer
		
		'Server port (typically 25)
		objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		
		
		objMessage.Configuration.Fields.Update
		objMessage.Send

	Set objMessage = Nothing
'**************************************************************************************
END SUB
