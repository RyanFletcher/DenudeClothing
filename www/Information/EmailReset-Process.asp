<% EmailAddressVar = Request.Form("EmailAddress")
	
	SourceVar = "SELECT Users.UserID, Users.FirstName, Users.Surname FROM Users WHERE EmailAddress = '" & EmailAddressVar 
		Call OpenRecordSet(UserInfo,1,SourceVar)
			UserIDVar = UserInfo.Field("UserID")
			FirstNameVar = UserInfo.Fields("Firstname")
			SurnameVar = UserInfo.Fields("Surname")
			
			EmailVar = "Hi" & FirstnameVar & " " & SurnameVar & ","
			EmailVar = EmailVar & "<p>You have recently requested you password to be reset. To reset your password, please click the link below</p>"
			EmailVar = EmailVar & "<a href='www.denudeclothing.com/Information/PasswordReset.asp?" & UserIDVar * "> Click here to reset your password </a>"
			EmailVar = EmailVar & "If you did not request a password request, please click here to report an issue with your account"
			
			
			