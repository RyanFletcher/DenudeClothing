<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
 <!--#include file="../../Connections/Connections.asp" -->
<%	
ItemIDVar = Request.QueryString("ItemID")
Set DeleteRecord = CreateObject("ADODB.Command")
		DeleteRecord.ActiveConnection = DEN_Conn(0)
		DeleteRecord.CommandText = "DELETE FROM Items WHERE ItemID ="& ItemIDVar
		DeleteRecord.Execute
		DeleteRecord.ActiveConnection.Close
        Response.Redirect("Items.asp")
 %>
