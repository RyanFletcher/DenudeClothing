<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--#include file="../Connections/Connections.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" href="../CSS/DenudeMain.css" type="text/css" />
<base target="_parent" />
<title>Accesories- iFrame</title>
</head>

<body>
           <div align="center">
           <table border="0" cellpadding="0" cellspacing="3">
           <tr align="center">
           <%
		     SourceVar = "SELECT ItemID,ItemTypeID,ItemName,UnitPrice,ImageSource,Available FROM Items WHERE ItemTypeID = 3"
			 	Call OpenRecordSet(Items,0,SourceVar)
					 While NOT Items.EOF
					 ItemIDVar = Items.Fields("ItemID")
					 ItemTypeIDVar = Items.Fields("ItemTypeID")
					 ItemNameVar = Items.Fields("ItemName")
					 UnitPriceVar = Items.Fields("UnitPrice")
					 ImageSourceVar = Items.Fields("ImageSource")
					 AvailableVar = Items.Fields("Available")
					 %>
           <form name="Item<%= ItemIDVar %>"  action="Checkout/Basket-Add-Process.asp?ItemID=<%= ItemIDVar %>&amp;ItemName=<%= ItemNameVar %>&amp;UnitPrice=<%= UnitPriceVar %>" method="post">
            <td width="200"><p id="FormLabel"><% If AvailableVar = True Then 
			response.Write(ItemNameVar)
			Else
			response.Write("Out Of Stock")
			End If %>
			</p><img width="170" height="170" src="../Images/<%= ImageSourceVar %>" name="<%= ItemNameVar %>" />
            <p id="FormLabel">£<%= UnitPriceVar %></p>
            <% If AvailableVar = True Then %>
            <input type="submit" value="Add to basket" />
            <% Else
			End If %>
            </td>
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
