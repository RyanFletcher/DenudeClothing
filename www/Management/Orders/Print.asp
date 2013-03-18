<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Dènudè Clothing</title>
<!--#include file="../../Connections/Connections.asp" -->
<% OrderIDVar =  Request.QueryString("OrderID") 

	SourceVar = "SELECT * FROM Orders INNER JOIN Users ON Orders.UserID = Users.UserID INNER JOIN OrderItems ON Orders.OrderID = OrderItems.OrderID"
		Call OpenRecordSet(GetOrder,0,SourceVar) 
				OrderedByVar = GetOrder.Fields("FirstName") & " " & GetOrder.Fields("Surname") 
				AddressVar = GetOrder.Fields("AddressLine1") & "<br />" & GetOrder.Fields("AddressLine2") & "<br />" & GetOrder.Fields("AddressLine2") & "<br />" & GetOrder.Fields("Postcode")
				SortCodeVar = GetOrder.Fields("SortCode")
				%>
				
  <link rel="stylesheet" href="../../CSS/bootstrap.css" type="text/css" /> 
  <link rel="stylesheet" href="../../CSS/DenudeMain.css" type="text/css" />
  <script src="../JavaScript/FormValidation.js" language="javascript"></script>
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
</head>
<body>
<table style="border-right: 1px solid #CCC; border-left: 1px solid #CCC;" width="945" height="100%" align="center">
 <tr>
  <td height="486" valign="top">
   <table border="0" width="945" align="top"> 
     <tr>
      <td width="315"></td>
      <td width="315" align="center"><p><img src="../../Images/DenudeLogo2.png" height="98" width="222" /></p></td>
      <td width="315">
      <table>
      <tr>
      	<td align="right"  width="300"><p>Invoice to: <%= OrderedByVar %></p>
        	<p>Invoice date: <%= now() %></p>
            <p>Address: <%= AddressVar %></p>
            <p>Sort Code: <%= SortCodeVar %></p>
            </td>
      </tr>
      </table>
      </td>
      </tr>
     </table>
     </td>
    </tr>
   <tr height="30" width="100%">
   <td>
   	<table>
     <tr>	
    <td width="315" height="73" align="center" valign="middle" id="Footer">Denude Clothing</td>
    <td width="315" valign="middle" align="center" id="Footer">Report Created By <%= Session("FirstName") & " " & Session("Surname") %></td>
    <td width="315" valign="middle" align="center" id="Footer"><%= now() %></td>
    </tr>
   </table>
   </td>
  </tr>  
  </table>
 </td>
</tr>
</table>
</body>
</html>
