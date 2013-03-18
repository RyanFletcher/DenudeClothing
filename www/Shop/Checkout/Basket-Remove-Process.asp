<!--#include file="../../Denude_Global_Functions.asp" -->
<!--#include file="../../Connections/Connections.asp" -->
<% ItemIDVar = Request.QueryString("ItemID")
   UserIDVar = Session("UserID")
   GuestIDVar = Session("GuestID")
   RedirectVar = Request.QueryString("Redirect")
   
Set InsertRecord = CreateObject("ADODB.Command")
    InsertRecord.ActiveConnection = DEN_Conn(0)
	If UserIDVar <> 0 Then
	DeleteVar = "DELETE FROM Basket WHERE ItemID =" & ItemIDVar & " AND UserID =" & UserIDVar
	Else
	DeleteVar = "DELETE FROM Basket WHERE ItemID =" & ItemIDVar & " AND GuestID =" & GuestIDVar
	End If
	InsertRecord.CommandText = DeleteVar
    InsertRecord.Execute
	InsertRecord.ActiveConnection.Close
		
	Response.Redirect(""& RedirectVar &".asp") %>
	