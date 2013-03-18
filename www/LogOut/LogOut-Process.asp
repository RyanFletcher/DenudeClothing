<!--#include file="../Connections/Connections.asp" -->
<!--#include file="../Denude_Global_Functions.asp" -->
<%
        UserIDVar = Session("UserID")
		Set DeleteRecord = CreateObject("ADODB.Command")
		DeleteRecord.ActiveConnection = DEN_Conn(0)
		DeleteRecord.CommandText = "DELETE FROM Basket WHERE UserID = " & UserIDVar
		DeleteRecord.Execute
		DeleteRecord.ActiveConnection.Close
		Session.Abandon()
		Response.Redirect("../Index.asp")
%>