<!--#include file="../../Connections/Connections.asp" -->
<!--#include file="../../Denude_Global_Functions.asp" -->
<% 

	UserIDVar = Session("UserID")
	
	SourceVar = "SELECT * FROM Users WHERE UserID=" & UserIDVar
		Call OpenRecordSet(UserInfo,0,SourceVar)
				FirstNameVar = UserInfo.Fields("FirstName")
				SurnameVar = UserInfo.Fields("Surname")
				EmailVar = UserInfo.Fields("EmailAddress")
				AddressLine1Var = UserInfo.Fields("AddressLine1")
				AddressLine2Var = UserInfo.Fields("AddressLine2")
				AddressLine3Var = UserInfo.Fields("AddressLine3")
				PostcodeVar = UserInfo.Fields("Postcode")
				CountyVar = UserInfo.Fields("County")
				CountryVar = UserInfo.Fields("Country")
				CardNumberVar = UserInfo.Fields("CardNumber")
				SecurityCodeVar = UserInfo.Fields("SecurityCode")
				AccountNumberVar = UserInfo.Fields("AccountNumber")
				SortCodeVar = UserInfo.Fields("SortCode")
				UserInfo.Close()
								
				Set InsertRecord = CreateObject("ADODB.Command")
					InsertRecord.ActiveConnection = DEN_Conn(0)  
					InsertRecord.CommandText = "INSERT INTO Orders (UserID, DateOrdered,GrossAmount, OrderStatusID) VALUES (" & UserIDVar & ",'" & now & "','0.00',1)"
					InsertRecord.Execute
					InsertRecord.ActiveConnection.Close
	
	SourceVar = "SELECT TOP (1) OrderID FROM Orders ORDER BY OrderID DESC"
		Call OpenRecordSet(GetOrder,0,SourceVar)
				OrderIDVar = GetOrder.Fields("OrderID")
				GetOrder.Close()
	
	SourceVar = "SELECT * FROM Basket WHERE UserID=" & UserIDVar
		Call OpenRecordSet(CollectItems,0,SourceVar)
			While NOT CollectItems.EOF
				ItemIDVar = CollectItems.Fields("ItemID")
				ItemNameVar = CollectItems.Fields("ItemName")
				ItemUnitPriceVar = CollectItems.Fields("UnitPrice")
				QuantityVar = CollectItems.Fields("Quantity")
				
				'Insert Each Item Into the OrderItems Table
				Set InsertRecord = CreateObject("ADODB.Command")
					InsertRecord.ActiveConnection = DEN_Conn(0)  
					InsertRecord.CommandText = "INSERT INTO OrderItems (OrderID, ItemID, ItemName, UnitPrice, Quantity) VALUES (" & OrderIDVar & "," & ItemIDVar & ",'" & ItemNameVar & "','" & ItemUnitPriceVar & "'," & QuantityVar &")"
					InsertRecord.Execute
					InsertRecord.ActiveConnection.Close

				
				CollectItems.MoveNext()
			Wend
			SourceVar = "SELECT SUM(OrderItems.UnitPrice * OrderItems.Quantity) AS UnitAmount, SUM(OrderItems.UnitPrice) AS GrossAmount, OrderItems.OrderID FROM OrderItems WHERE OrderID=" & OrderIDVar & " GROUP BY OrderID"
		Call OpenRecordSet(GetOrderValue,0,SourceVar)
		 GrossAmountVar = GetOrderValue.Fields("GrossAmount")
		 	GetOrderValue.Close()
		CollectItems.Close()
		
				Set UpdateRecord = CreateObject("ADODB.Command")
					UpdateRecord.ActiveConnection = DEN_Conn(0)  
					UpdateRecord.CommandText = "UPDATE Orders SET GrossAmount ='" & GrossAmountVar & "' WHERE OrderID =" & OrderIDVar
					UpdateRecord.Execute
					UpdateRecord.ActiveConnection.Close
					
				Set DeleteRecord = CreateObject("ADODB.Command")
					DeleteRecord.ActiveConnection = DEN_Conn(0)
					DeleteRecord.CommandText = "DELETE FROM Basket WHERE UserID =" & UserIDVar
					DeleteRecord.Execute
					DeleteRecord.ActiveConnection.Close		
					
				Response.Redirect("Basket-Checkout-OrderSummary.asp")
				%>	
		
		

						
			