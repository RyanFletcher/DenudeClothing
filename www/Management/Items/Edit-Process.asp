<!--#include file="../../Connections/Connections.asp" --> 
<%  
	ItemIDVar = Request.Form("ItemID")
	ItemNameVar = Request.Form("ItemName")
	UnitPriceVar = Request.Form("UnitPrice")
	ImageSourceVar = Request.Form("ImageSource")
	ItemTypeIDVar = Request.Form("ItemTypeID")
	AvailableVar = Request.Form("Available")
	 
	 If len(AvailableVar) = 0 Then
	 	AvailableVar = 0
	 End If	
	 Set UpdateRecord = CreateObject("ADODB.Command")
		UpdateRecord.ActiveConnection = DEN_Conn(0)
		UpdateRecord.CommandText = "UPDATE Items SET ItemTypeID = " & ItemTypeIDVar & ", ItemName = '" & ItemNameVar & "', UnitPrice = '" & UnitPriceVar & "', ImageSource = '" & ImageSourceVar & "', Available = " & AvailableVar & " WHERE ItemID = " & ItemIDVar
		UpdateRecord.Execute
		UpdateRecord.ActiveConnection.Close
        Response.Redirect("Items.asp") %>
