<!--#include file="../../Connections/Connections.asp" -->
<!--#include file="../../Functions/PrivateFunctions.asp" -->
<% 'Login and authenticate Client/User
	
	UsernameVar = Request.Form("Username")
	PasswordVar = Request.Form("Password")
	If IsEmpty(UsernameVar) OR IsEmpty(PasswordVar) Then
		Response.Redirect("../SessionTimeOut.asp")
	Else
		SourceVar = "SELECT Username FROM Clients WHERE Username = '" & SQLTextString(UsernameVar) & "'"
		Call OpenRecordSet(AuthenticateUsername,1,SourceVar)
		If AuthenticateUsername.EOF Then
			Response.Redirect("Login.asp?InvalidUsername=1")
		Else
		 SourceVar = "SELECT Clients.ClientID, Clients.Username, Clients.Password FROM Clients WHERE Username = '" & SQLTextString(UsernameVar) & "' AND Password = '" & SQLTextString(PasswordVar) & "'"
			Call OpenRecordSet(AuthenticatePassword,1,SourceVar)
				If AuthenticatePassword.EOF Then
					Response.Redirect("Login.asp?InvalidPassword=1")
				Else
					ClientIDVar = AuthenticatePassword.Fields("ClientID")
					 SourceVar = "SELECT * FROM Clients WHERE ClientID = " & ClientIDVar
					 Call OpenRecordSet(ClientInfo,1,SourceVar)
					 	 Session("ClientID") = ClientInfo.Fields("ClientID")
						 Session("ClientName") = ClientInfo.Fields("ClientName")
						 Session("FirstName") = ClientInfo.Fields("FirstName")
						 Session("Surname") = ClientInfo.Fields("Surname")
						 Session("Email") = ClientInfo.Fields("Email")
						 Session("Telephone") = ClientInfo.Fields("Telephone")
						 Session("Address") = ClientInfo.Fields("Address1") & "<br />" & ClientInfo.Fields("Address2") & "<br />" & ClientInfo.Fields("Address3") & "<br />" & ClientInfo.Fields("Address4") & "<br />" & ClientInfo.Fields("Address5")
						 ClientInfo.Close()
						AuthenticatePassword.Close()
					AuthenticateUsername.Close()	 
				Response.Redirect("../Dashboard/Dashboard.asp")
			End If
		End If				
	End If				
%>				