<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%	
	Set Mail = Server.CreateObject("Persits.MailSender")
			Mail.Host = "mail.denudeclothing.com"
			Mail.Username = "sendmail@mailserver.denudeclothing.com"
			Mail.Password = "INDSendMail"
			Mail.From = "mail@denudeclothing.com"
			Mail.FromName = "Denude Clothing"
			Mail.Subject = "Ticket Status Alert: " & CompanyNameVar & ":" &  OutletNameVar &  " (" & TicketIDVar & ": " & TitleVar & ")"
			Mail.Body = EmailBodyVar
			Mail.AddAddress EmailAddressesArrayVar(0)
			
			'Other Recipients
			If uBound(EmailAddressesArrayVar) > 0 Then
				For EmailAddressNo = 1 to uBound(EmailAddressesArrayVar)
					Mail.AddCC EmailAddressesArrayVar(EmailAddressNo)
				Next
			End If

'			'Attachments
'			If FilePathVar <> "" Then													
'				FilePathArrayVar = Split(Left(FilePathVar,Len(FilePathVar)-1),",")													
'				For x = 0 to uBound(FilePathArrayVar)
'					Mail.AddAttachment Server.MapPath(FilePathArrayVar(x)) 					
'				Next								
'			End If
			
			Mail.IsHTML = True							
			Mail.Queue = True
			Mail.Send
		Set Mail = Nothing	
%>