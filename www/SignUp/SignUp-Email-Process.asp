<!--#include file="../Connections/Connections.asp" -->
  
 <% 
	      FirstNameVar = Request.Form("FirstName")
	      SurnameVar = Request.Form("Surname")
	      EmailAddressVar = Request.Form("EmailAddress")
	      PasswordVar = Request.Form("Password")
		  
		'Check if the email address already exists
		SourceVar = "SELECT * FROM Users WHERE EmailAddress = '" & EmailAddressVar & "'"
	Call OpenRecordSet(CheckDuplicatedUser,0,SourceVar)
			If NOT CheckDuplicatedUser.EOF Then
			Session.Abandon()
			Response.Redirect("../Index.asp?DuplicatedUser=1")
		    Else
			End If
		 CheckDuplicatedUser.Close()

	   
	  Set InsertRecord = CreateObject("ADODB.Command")
		  InsertRecord.ActiveConnection = DEN_Conn(0)
		  InsertRecord.CommandText = "INSERT INTO Users (FirstName,Surname,EmailAddress,DenudeUser) VALUES ('" & FirstNameVar & "','" & SurnameVar & "','" & EmailAddressVar & "',0)"
		  InsertRecord.Execute
	 	  InsertRecord.ActiveConnection.Close
		
	SourceVar = "SELECT * FROM Users WHERE EmailAddress = '" & EmailAddressVar & "'"
	Call OpenRecordSet(Authenticate,0,SourceVar)
		  Session("UserID") = Authenticate.Fields("UserID")
		  Session("FirstName") = Authenticate.Fields("FirstName")
		  Session("Surname") = Authenticate.Fields("Surname")
		  Session("EmailAddress") = Authenticate.Fields("EmailAddress")
		  Session("DenudeUser") = Authenticate.Fields("DenudeUser")
		 Authenticate.Close()
	Response.Redirect("SignUp-Email-Confirmation.asp")		
 %>