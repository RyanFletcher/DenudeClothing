<!--#include file="../../Connections/Connections.asp" -->
<% UserIDVar = Request.Form("UserID")

    Set DeleteRecord = CreateObject("ADODB.Command")
		DeleteRecord.ActiveConnection = DEN_Conn(0)
		DeleteRecord.CommandText = "DELETE FROM Users WHERE UserID=" & UserIDVar 
		DeleteRecord.Execute
		DeleteRecord.ActiveConnection.Close
        Response.Redirect("Users.asp") %>