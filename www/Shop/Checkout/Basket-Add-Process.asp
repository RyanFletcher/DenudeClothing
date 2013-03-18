<!--#include file="../../Denude_Global_Functions.asp" -->
<!--#include file="../../Connections/Connections.asp" -->
<% 
   If Session("UserID") = "" Then
   		UserIDVar = 0
   Else UserIDVar = Session("UserID")
   End If		
   If Session("GuestID") = "" Then
   		GuestIDVar = 0
   Else GuestIDVar = Session("GuestID")
   End If
   ItemIDVar = Request.QueryString("ItemID")
   ItemNameVar = Request.QueryString("ItemName")
   ItemUnitPriceVar = Request.QueryString("UnitPrice")
   
  'Determine whether the item already exists
   If GuestIDVar = 0 Then 
   SourceVar = "SELECT * FROM Basket WHERE UserID =" & UserIDVar &" AND ItemID=" & ItemIDVar & " AND Quantity >= 1"
   Else
   SourceVar = "SELECT * FROM Basket WHERE GuestID =" & GuestIDVar &" AND ItemID=" & ItemIDVar & " AND Quantity >= 1"
   End If
   	Call OpenRecordSet(CheckBasket,0,SourceVar)
		If CheckBasket.EOF Then
	Set InsertRecord = CreateObject("ADODB.Command")
		InsertRecord.ActiveConnection = DEN_Conn(0)  
		InsertRecord.CommandText = "INSERT INTO Basket (UserID, ItemID, ItemName, UnitPrice, Quantity, GuestID) VALUES (" & UserIDVar & "," & ItemIDVar & ",'" & ItemNameVar & "','" & ItemUnitPriceVar & "',1," & GuestIDVar & ")"
		InsertRecord.Execute
		InsertRecord.ActiveConnection.Close
	Else
		If GuestIDVar = 0 Then
   		SourceVar = "SELECT * FROM Basket WHERE UserID =" & UserIDVar &" AND ItemID=" & ItemIDVar
		Else
   		SourceVar = "SELECT * FROM Basket WHERE GuestID =" & GuestIDVar &" AND ItemID=" & ItemIDVar
		End If
   			Call OpenRecordSet(AddQuantity,0,SourceVar)
				QuantityVar = AddQuantity.Fields("Quantity")
				QuantityVar = QuantityVar + 1
			  AddQuantity.Close()
		If GuestIDVar = 0 Then
		UpdateVar = "UPDATE Basket SET Quantity =" & QuantityVar & " WHERE UserID =" & UserIDVar &" AND ItemID=" & ItemIDVar	 
		Else
		UpdateVar = "UPDATE Basket SET Quantity =" & QuantityVar & " WHERE GuestID =" & GuestIDVar &" AND ItemID=" & ItemIDVar 
		End If
	Set UpdateRecord = CreateObject("ADODB.Command")
		UpdateRecord.ActiveConnection = DEN_Conn(0)
		UpdateRecord.CommandText = UpdateVar
		UpdateRecord.Execute
		UpdateRecord.ActiveConnection.Close
		
	End If
		Response.Redirect("Basket.asp") 
		%>
	