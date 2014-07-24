'TODO: Clean this up and refactor to a SUB


'Create the message
reportMessage = msg1 & chr(13) & chr(10) & msg2 & chr(13) & chr(10)

	'Email message
	Dim objMessage, strTo, strFrom, strSubject, strBody
	Const SMTPServer = "smtp.gmail.com"
	Const SMTPLogon = "your-email@gmail.com"
	Const SMTPPassword = "your-password"
	Const SMTPSSL = True
	Const SMTPPort = 465

	Const cdoSendUsingPickup = 1 	'Send message using local SMTP service pickup directory.
	Const cdoSendUsingPort = 2 	'Send the message using SMTP over TCP/IP networking.

	Const cdoAnonymous = 0 	' No authentication
	Const cdoBasic = 1 	' BASIC clear text authentication
	Const cdoNTLM = 2 	' NTLM, Microsoft proprietary authentication
	Set objMessage= CreateObject("CDO.Message")            
									
	strTo="to@some-email.com"
	strFrom="from@some-email.com"
	strSubject="Message Subject" & emailtimestamp
	strBody=reportMessage & " " & emailtimestamp

		objMessage.To = strTo
		objMessage.From = strFrom
		objMessage.Subject = strSubject
		objMessage.TextBody = strBody
		objMessage.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2	

		objMessage.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTPServer

		objMessage.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic

		objMessage.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/sendusername") = SMTPLogon

		objMessage.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/sendpassword") = SMTPPassword

		objMessage.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTPPort

		objMessage.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = SMTPSSL

		objMessage.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
		
		objMessage.Configuration.Fields.Update
		objMessage.Send

	Set objMessage = Nothing
'msgbox "Email sent!"

