<!--#include file="../../Connections/Connections.asp" -->
<% 
   UserIDVar = Request.QueryString("UserID")
   ItemIDVar = Request.QueryString("ItemID")
   ItemNameVar = Request.QueryString("ItemName")
   ItemUnitPriceVar = Request.QueryString("UnitPrice")
   RedirectVar = Request.QueryString("Redirect")
   
  'Determine whether the user already exists 
  If UserIDVar <> "" Then
   
   SourceVar = "SELECT * FROM Basket WHERE UserID =" & UserIDVar &" AND ItemID=" & ItemIDVar & " AND Quantity >= 1"
   	Call OpenRecordSet(CheckBasket,0,SourceVar)
		If CheckBasket.EOF Then
	Set InsertRecord = CreateObject("ADODB.Command")
		InsertRecord.ActiveConnection = DEN_Conn(0)  
		InsertRecord.CommandText = "INSERT INTO Basket (UserID, ItemID, ItemName, UnitPrice, Quantity, GuestID) VALUES (" & UserIDVar & "," & ItemIDVar & ",'" & ItemNameVar & "','" & ItemUnitPriceVar & "',1,0)"
		InsertRecord.Execute
		InsertRecord.ActiveConnection.Close
	Else
   		SourceVar = "SELECT * FROM Basket WHERE UserID =" & UserIDVar &" AND ItemID=" & ItemIDVar
   			Call OpenRecordSet(AddQuantity,0,SourceVar)
				QuantityVar = AddQuantity.Fields("Quantity")
				QuantityVar = QuantityVar + 1
			  AddQuantity.Close()
	Set UpdateRecord = CreateObject("ADODB.Command")
		UpdateRecord.ActiveConnection = DEN_Conn(0)  
		UpdateRecord.CommandText = "UPDATE Basket SET Quantity =" & QuantityVar & " WHERE UserID =" & UserIDVar &" AND ItemID=" & ItemIDVar
		UpdateRecord.Execute
		UpdateRecord.ActiveConnection.Close
		
	End If

		Response.Redirect("Basket.asp") 
		
'If the user does NOT exist then redirect to the sign up page but carry the item id		
Else Response.Redirect("../../SignUp/SignUp.asp?ItemID=" & ItemIDVar)

End If
		%>
	