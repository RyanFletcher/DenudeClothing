<!--#include file="../Connections/Connections.asp" -->
  
 <% 
	   'Gather User Information	
	   FirstNameVar = Request.Form("FirstName")
	   SurnameVar = Request.Form("Surname")
	   EmailAddressVar = Request.Form("EmailAddress")
	   PasswordVar = Request.Form("Password")
	   ItemIDVar = Request.Form("ItemID")
	   
	    'Log user in and update session information
		SourceVar = "SELECT * FROM Users WHERE EmailAddress = '" & EmailAddressVar & "'"
	Call OpenRecordSet(CheckDuplicatedUser,0,SourceVar)
			If NOT CheckDuplicatedUser.EOF Then
			Response.Redirect("../Index.asp?DuplicatedUser=1")
		    Else
			End If
		 Authenticate.Close()

		'Insert New User Record
	Set InsertRecord = CreateObject("ADODB.Command")
		InsertRecord.ActiveConnection = DEN_Conn(0)
		InsertRecord.CommandText = "INSERT INTO Users (FirstName,Surname,EmailAddress,Password,DenudeUser) VALUES ('" & FirstNameVar & "','" & SurnameVar & "','" & EmailAddressVar & "','" & PasswordVar & "',0)"
		InsertRecord.Execute
		InsertRecord.ActiveConnection.Close
		
		'Log user in and update session information
		SourceVar = "SELECT * FROM Users WHERE EmailAddress = '" & EmailAddressVar & "' AND Password = '" & PasswordVar & "'"
	Call OpenRecordSet(Authenticate,0,SourceVar)
		 If Authenticate.EOF Then
		 Response.Redirect("../Index.asp")
		 Else Session("UserID") = Authenticate.Fields("UserID")
		 Session("FirstName") = Authenticate.Fields("FirstName")
		 Session("Surname") = Authenticate.Fields("Surname")
		 Session("EmailAddress") = Authenticate.Fields("EmailAddress")
		 Session("DenudeUser") = Authenticate.Fields("DenudeUser")
		 End If
		 Authenticate.Close()
		 
	'Insert Record into basket if the user did not exsist when adding an item to the basket	 
	If ItemIDVar > 0 Then
		Set InsertRecord = CreateObject("ADODB.Command")
			InsertRecord.ActiveConnection =  DEN_Conn(0)
			InsertRecord.CommandText = "INSERT INTO Basket (UserID, ItemID, Quantity, GuestID) VALUES (" & Session("UserID") & "," & ItemIDVar & ",1,0)"
			InsertRecord.Execute
			InsertRecord.ActiveConnection.Close
			Else
	End If
		Response.Redirect("SignUp-Confirmation.asp?ItemID=" & ItemIDVar)		
 %>