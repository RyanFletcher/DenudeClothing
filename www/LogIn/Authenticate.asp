<!--#include file="../Connections/Connections.asp" -->
<!--#include file="../Denude_Global_Functions.asp" -->
<%	
	EmailAddressVar = Request.Form("EmailAddress")
    PasswordVar = Request.Form("Password") 	
	SourceVar = "SELECT EmailAddress FROM Users WHERE EmailAddress = '" & EmailAddressVar & "'"
		 Call OpenRecordSet(Authenticate,0,SourceVar)
		 	  If Authenticate.EOF Then
			  Response.Redirect("../Index.asp?Credentials=1")
			  Else
			  EmailAddressVar = Authenticate.Fields("EmailAddress")
			  Authenticate.close()
			  End If
			  
    SourceVar = "SELECT UserID,FirstName,Surname,EmailAddress,Password,DenudeUser FROM Users WHERE EmailAddress ='" & EmailAddressVar & "' AND Password='" & PasswordVar & "'"          
	Call OpenRecordSet(Authenticate,0,SourceVar)
		 If Authenticate.EOF Then
		 Response.Redirect("../Index.asp?Credentials=1")
		 Else Session("UserID") = Authenticate.Fields("UserID")
		 Session("FirstName") = Authenticate.Fields("FirstName")
		 Session("Surname") = Authenticate.Fields("Surname")
		 Session("EmailAddress") = Authenticate.Fields("EmailAddress")
		 Session("Password") = Authenticate.Fields("Password")
		 Session("DenudeUser") = Authenticate.Fields("DenudeUser")
		 End If
		 Authenticate.Close()
		 
		 Response.Redirect("/Index.asp") 
%>
		 