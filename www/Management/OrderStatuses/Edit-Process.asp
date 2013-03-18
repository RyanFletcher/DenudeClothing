<!--#include file="../../Connections/Connections.asp" --> 
<%  OrderStatusIDVar = Request.Form("OrderStatusID")
	OrderStatusVar = Request.Form("OrderStatus")
	EmailAddressVar = Request.Form("EmailAddress")

	
    Set UpdateRecord = CreateObject("ADODB.Command")
		UpdateRecord.ActiveConnection = DEN_Conn(0)
		UpdateRecord.CommandText = "UPDATE OrderStatuses SET OrderStatus = '" & OrderStatusVar & "', EmailAddressRedirection = '" & EmailAddressVar & "' WHERE OrderStatusID = " & OrderStatusIDVar
		UpdateRecord.Execute
		UpdateRecord.ActiveConnection.Close
        Response.Redirect("OrderStatuses.asp") %>