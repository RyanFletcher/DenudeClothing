<!--#include file="../../Connections/Connections.asp" -->
<% 
   UserIDVar = Request.Form("UserID")
   FirstNameVar = Request.Form("FirstName")
   SurnameVar = Request.Form("Surname")
   EmailAddressVar = Request.Form("EmailAddress")
   PasswordVar = Request.Form("Password")
   
    Set UpdateRecord = CreateObject("ADODB.Command")
		UpdateRecord.ActiveConnection = DEN_Conn(0)
		UpdateRecord.CommandText = "UPDATE Users SET FirstName ='" & FirstNameVar & "', Surname ='" & SurnameVar & "', EmailAddress ='" & EmailAddressVar & "', Password ='" & PasswordVar & "' WHERE UserID =" & UserIDVar
		UpdateRecord.Execute
		UpdateRecord.ActiveConnection.Close
        Response.Redirect("Users.asp") %>