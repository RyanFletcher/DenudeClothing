<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="../Connections/Connections.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="../CSS/DenudeMain.css" type="text/css" />
<base target="_parent" />
<title>Mens- iFrame</title>
</head>

<body>
           <div align="center">
           <table border="0" cellpadding="0" cellspacing="3">
           <tr align="center">
           <%
		   	 'Cycle through Items With an item type of 1 (Mens)
		     SourceVar = "SELECT ItemID,ItemTypeID,ItemName,UnitPrice,ImageSource FROM Items WHERE ItemTypeID = 1"
			 	Call OpenRecordSet(Items,0,SourceVar)
					'For each item, set variables and create a form
					 While NOT Items.EOF
					 ItemIDVar = Items.Fields("ItemID")
					 ItemTypeIDVar = Items.Fields("ItemTypeID")
					 ItemNameVar = Items.Fields("ItemName")
					 UnitPriceVar = Items.Fields("UnitPrice")
					 ImageSourceVar = Items.Fields("ImageSource")
					 %>
           <form name="Item<%= ItemIDVar %>"  action="Checkout/Basket-Add-Process.asp?ItemID=<%= ItemIDVar %>&amp;ItemName=<%= ItemNameVar %>&amp;UnitPrice=<%= UnitPriceVar %>" method="post">
            <td width="200"><p id="FormLabel"><%= ItemNameVar %></p><img width="170" height="170" src="../Images/<%= ImageSourceVar %>" name="<%= ItemNameVar %>" />
            <p id="FormLabel">£<%= UnitPriceVar %></p>
            <input type="submit" value="Add to basket" /></td>
           </form>
           
           <% Items.MoveNext()
		     Wend
		   Items.Close() %>
           </tr>
           </table>
            </div>
          </td>
         </tr>
        </table>
       </div>
</body>
</html>
