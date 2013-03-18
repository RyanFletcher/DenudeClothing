<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><!-- InstanceBegin template="/Templates/DenudeTemplate.dwt.asp" codeOutsideHTMLIsLocked="false" --><head>
 <% WrongCredentialsVar = Request.QueryString("Credentials") %>
 <!-- InstanceBeginEditable name="Title" -->
 <title>Dènudè Clothing</title>
 <!-- InstanceEndEditable --> 
 <!--#include file="../../Connections/Connections.asp" -->
 <!--#include file="../../Denude_Global_Functions.asp" -->
 <link rel="stylesheet" href="../../CSS/bootstrap.css" type="text/css" /> 
 <link rel="stylesheet" href="../../CSS/DenudeMain.css" type="text/css" />
 <link rel="icon"type="image/png" href="../../Images/DenudeLogo2.png">
 <script src="../../JavaScript/FormValidation.js" language="javascript"></script>
 <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
</head>
<body> 
<table cellpadding="50" width="100%" height="100%" align="center">
 <tr valign="top">
  <td width="10">&nbsp;</td>
   <td width="100" height="100%" valign="top">
    <table border="0" width="100%" align="center"> 
     <tr> 
      <td style="border-bottom: 1px solid #CCC;" height="100%" valign="top">
       <div align="right">  
	    <% If Session("UserID") ="" AND Session("GuestID") = "" Then %>
         <p id="FormLabel">
          <form class="FormLabel" name="Login" id="Login" method="post" action="../../../Login/Authenticate.asp">
            <% If WrongCredentialsVar > 0 Then %>
            <p id="FormLabel"> Incorrect Credentials </p>
            <% End If %>
            <input placeholder="Email Address" name="EmailAddress" width="10" type="text" /> <input width="10" placeholder="Password" name="Password" type="password" /> <input type="submit" value="Log In" class="Button"><br />
           <a href="Information-Credentials.asp">Forgotten your username or password?</a>
         </form>
        <% Else 
			If Session("UserID") <> 0 Then
		SourceVar = "SELECT SUM(Items.UnitPrice * Basket.Quantity) AS Total, SUM(Basket.Quantity) AS Quantity,Basket.UserID FROM Basket INNER JOIN Items ON Basket.ItemID = Items.ItemID WHERE UserID = " & Session("UserID") & " GROUP BY Basket.UserID"
			Else
		SourceVar = "SELECT SUM(Items.UnitPrice * Basket.Quantity) AS Total, SUM(Basket.Quantity) AS Quantity,Basket.GuestID FROM Basket INNER JOIN Items ON Basket.ItemID = Items.ItemID WHERE GuestID = " & Session("GuestID") & " GROUP BY Basket.GuestID"
			End If
		    Call OpenRecordSet(BasketInfo,0,SourceVar)
		      While NOT BasketInfo.EOF
			  TotalOrderValueVar = BasketInfo.Fields("Total")
			  QuantityVar = BasketInfo.Fields("Quantity")
			  BasketInfo.MoveNext()
			  Wend
		      BasketInfo.Close() %>
             <form name="LogOut" id="LogOut" action="../../LogOut/LogOut-Process.asp">Logged in as <% If Session("GuestID") <> "" Then Response.Write("a Guest") Else Response.Write(Session("FirstName")) & " " & Session("Surname") End If %><% If Session("GuestID") <> "" Then
			     Else %>  (<a href="../../Profile/index.asp">Edit Profile</a>)<% End If %><br />
                  <a href="../../Shop/Checkout/Basket.asp">Your Basket</a>: £<% If TotalOrderValueVar = "" Then %>0.00<% Else %><%= TotalOrderValueVar %>.00
                 <% End If %> (<% If QuantityVar = 0 Then %>0 Items<% Else %>
	            <%= QuantityVar %> Item<% If RecordNumberVar <> 1 then %>s<% Else 
	           End If 
	          End If %>)<br />
             <% If Session("UserID") <> "" Then %>
            <input type="submit" value="Log Out" class="Button">
           <% Else
		End If %>
            </form>
          <% End If %> 
        </div>
       </td>
      </tr>
     <tr>
      <td width="100%">
       <table cellpadding="0" cellspacing="0" border="0" width="100%" align="bottom">
        <tr>
         <td height="195" colspan="6" align="right"><img src="../../Images/Denude-Cropped.jpg" height="195" width="360" /></td>
        </tr>
        </table>
       <table cellpadding="15" cellspacing="0" border="0" width="100%" align="center">
        <tr id="Menu" align="center">
            <td id="Menu" style="border-bottom: 1px solid #CCC;"><a href="../../Index.asp">HOME</a></td>
            <td id="Menu" style="border-bottom: 1px solid #CCC; text-decoration:line-through;">ABOUT</td>
            <td id="Menu" style="border-bottom: 1px solid #CCC; text-decoration:line-through;">SHOP</td>
            <td id="Menu" style="border-bottom: 1px solid #CCC; text-decoration:line-through;">CONTACT</td>
            <td id="Menu" style="border-bottom: 1px solid #CCC; text-decoration:line-through;">HELP</td>
            <% If Session("DenudeUser") = "True" Then %>
            <td id="Menu" style="border-bottom: 1px solid #CCC; "><a href="../index.asp">MANAGEMENT</a></td>
        	<% End If %>
       </tr>
      </table>
     </td>
    </tr>
    <tr>
    <td> 
     <table cellpadding="0" cellspacing="0" border="0" width="100%" align="center">
      <tr>
       <td colspan="6">
  	    <!-- InstanceBeginEditable name="MainContent" -->
        <table cellpadding="15" cellspacing="0" border="0" width="930" align="center">
        <tr><td colspan="6" align="right"><a href="Add.asp">Add an Item</a></td></tr>
         <tr id="Menu" style="border-top:none;" align="center">      	   <td width="338" id="Menu" style="border-bottom: 1px solid #CCC;"><a href="../Users/index.asp">USERS</a></td>
           <td width="137" id="Menu" style="border-bottom: 1px solid #CCC;"><a href="../Orders/index.asp">ORDERS</a></td>
           <td width="365" id="Menu" style="border-bottom: 1px solid #CCC;"><a href="../OrderStatuses/index.asp">ORDER STATUSES</a></td>
           <td width="365" id="Menu" style="border-bottom: 1px solid #CCC;"><a href="index.asp">ITEMS</a></td>
</tr>
       </table>
           <table border="0" width="100%" height="86">
           <tr>
           <td height="28">ItemID</td><td>Item Name</td><td>Price</td><td>Item Type</td>
           </tr>
           <% SourceVar = "SELECT * FROM Items INNER JOIN ItemTypes ON Items.ItemTypeID = ItemTypes.ItemTypeID"
		   		Call OpenRecordSet(GetItems,0,SourceVar)
					If GetItems.EOF Then
					%>x
                    <tr>
                    <td height="20" colspan="6">No Current Items</td>
                    </tr>
                    <% Else 
					While NOT GetItems.EOF
						ItemIDVar = GetItems.Fields("ItemID")
						ItemNameVar = GetItems.Fields("ItemName")
						ItemPriceVar = GetItems.Fields("UnitPrice")
						ItemTypeVar = GetItems.Fields("ItemType")
						%>
            <tr style="border-bottom: sold 1px #000;">
             <td height="30"><a href="Edit.asp?ItemID=<%= ItemIDVar %>"><%= ItemIDVar %></a></td><td><%= ItemNameVar %></td><td><%= ItemPriceVar %></td><td><%= ItemTypeVar %>
            </tr>
            
            <% GetItems.MoveNext()
			Wend
		GetItems.Close()
	End If
		%> 
           </table>
        
 <title>Dènudè Clothing</title>
 <!-- InstanceEndEditable -->
       </td>
      </tr>  
     </table>
    </td>
   </tr>
   <tr>
    <td>
    </td>
   </tr>
  </table>
  <table width="100%" align="center">
   <tr height="30" width="100%">
    <td width="315"  valign="middle" id="Footer">
          <p align="left"><a href="https://twitter.com/DenudeClothing" class="twitter-follow-button" data-show-count="false">Follow @DenudeClothing</a>
<script>!function(d,s,id){var js,fjs=d.getElementsByTagName(s)[0];if(!d.getElementById(id)){js=d.createElement(s);js.id=id;js.src="//platform.twitter.com/widgets.js";fjs.parentNode.insertBefore(js,fjs);}}(document,"script","twitter-wjs");</script>
          <p align="left">
            <style>
          .ig-b- { display: inline-block; }
.ig-b- img { visibility: hidden; }
.ig-b-:hover { background-position: 0 -60px; } .ig-b-:active { background-position: 0 -120px; }
.ig-b-v-24 { width: 137px; height: 24px; background: url(//badges.instagram.com/static/images/ig-badge-view-sprite-24.png) no-repeat 0 0; }
@media only screen and (-webkit-min-device-pixel-ratio: 2), only screen and (min--moz-device-pixel-ratio: 2), only screen and (-o-min-device-pixel-ratio: 2 / 1), only screen and (min-device-pixel-ratio: 2), only screen and (min-resolution: 192dpi), only screen and (min-resolution: 2dppx) {
.ig-b-v-24 { background-image: url(//badges.instagram.com/static/images/ig-badge-view-sprite-24@2x.png); background-size: 160px 178px; }</style>
            <a href="http://instagram.com/dnde_clothing?ref=badge" target="_blank" class="ig-b- ig-b-v-24"><img src="//badges.instagram.com/static/images/ig-badge-view-24.png" alt="Instagram" /></a>
          <p align="left"></td>
    <td width="315"  valign="middle" id="Footer">
      <p align="center">Dènudè Ltd | 2012 | Copyright ©</p>
      <p align="center"><a href="Information/FAQs.asp">FAQ's</a>|<a href="Information/TermsAndConditions.asp">Terms &amp; Conditions</a>|<a href="Information/Contact.asp">Contact</a>|<a href="Information/Help.asp">Help</a></p>
    </td>
    <td width="315"  valign="middle" id="Footer">
     <div class="progress progress-striped active">
     <div class="bar" style="width: 75%;">75% Complete</div>
     </div>
    </td>
    </tr>
  </table>
 </td>
 <td width="10">&nbsp;</td>
</tr>
</table>
</body>
<!-- InstanceEnd --></html>