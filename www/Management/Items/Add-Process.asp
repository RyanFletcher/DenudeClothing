<!--#include file="../../Connections/Connections.asp" --> 
<% 

	ItemNameVar = Request.Form("ItemName")
	ItemPriceVar = Request.Form("UnitPrice")
	ItemImageSourceVar = Request.Form("ImageSource")
	ItemTypeIDVar = Request.Form("ItemTypeID")
	If Request.Form("Available") = "" Then
		AvailableVar = 1
	Else
	End If	
	

	Set InsetRecord = CreateObject("ADODB.Command")
		InsetRecord.ActiveConnection = DEN_Conn(0)
		InsetRecord.CommandText = "INSERT INTO Items (ItemName, UnitPrice, ImageSource, ItemTypeID, Available) VALUES ('"& ItemNameVar &"','" & UnitPriceVar & "','" & ItemImageSourceVar & "'," & ItemTypeIDVar & "," & AvailableVar & ")"
		InsetRecord.Execute
		InsetRecord.ActiveConnection.Close
        Response.Redirect("Items.asp") %>
