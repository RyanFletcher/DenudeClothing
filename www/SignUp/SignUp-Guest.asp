<!--#include file="../Connections/Connections.asp" -->

<% 	   FirstNameVar = Request.Form("FirstName")
	   SurnameVar = Request.Form("Surname")
	   AddressVar = Request.Form("Address")
	   EmailAddressVar = Request.Form("EmailAddress")
	   ItemIDVar = Request.QueryString("ItemID")
		
				Set InsertRecord = CreateObject("ADODB.Command")
					InsertRecord.ActiveConnection = DEN_Conn(0)
					InsertRecord.CommandText = "INSERT INTO Guests (FirstName, Surname, Address, EmailAddress) VALUES ('" & FirstNameVar & "','" & SurnameVar & "','" & AddressVar & "','" & EmailAddressVar & "')"
					InsertRecord.Execute()
					InsertRecord.ActiveConnection.Close()
					
				SourceVar = "SELECT TOP (1) GuestID, FirstName, Surname, Address, EmailAddress FROM Guests ORDER BY GuestID DESC"
					Call OpenRecordSet(GetGuest,0,SourceVar)
							Session("GuestID") = GetGuest.Fields("GuestID")
							GuestIDVar = GetGuest.Fields("GuestID")
							GetGuest.Close()
							
			If ItemIDVar > 0 Then
		Set InsertRecord = CreateObject("ADODB.Command")
			InsertRecord.ActiveConnection =  DEN_Conn(0)
			InsertRecord.CommandText = "INSERT INTO Basket (UserID, ItemID, Quantity, GuestID) VALUES (0," & ItemIDVar & ",1," & GuestIDVar & ")"
			InsertRecord.Execute
			InsertRecord.ActiveConnection.Close
		Response.Redirect("../Shop/Checkout/Basket.asp")		
			Else
		Response.Redirect("../Index.asp")		
	End If
					

				%>
