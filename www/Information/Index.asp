<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><!-- InstanceBegin template="/Templates/DenudeShopTemplate.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<!--#include file="../Connections/Connections.asp" -->
 <link rel="stylesheet" href="../CSS/DenudeMain.css" type="text/css" />
 <script src="../JavaScript/JQuery.js" type="text/javascript"></script>
 <script src="../JavaScript/Bootstrap.min.js" type="text/javascript"></script>
  <!-- InstanceBeginEditable name="doctitle" -->
<title>Dènudè Template</title>
<!-- InstanceEndEditable -->
  <!-- InstanceBeginEditable name="head" -->
<!-- InstanceEndEditable -->
</head> 
<body>
<table cellpadding="0" cellspacing="0" border="0" width="100%" align="center">
 <tr>
  <td>&nbsp;</td>
 </tr>
 <tr>
 <td id="HeaderLeft" valign="top" height="34" width="624"><a target="_blank" href="https://www.twitter.com/denudeclothing"><img src="../Images/twitter.jpeg" width="34" height="34" /></a></td>
  <td width="324" align="center" valign="top" id="HeaderRight">
    <div align="center" class="FormLabel">  
	  <%If Session("UserID") = "" then %>
       <p id="FormLabel">
        <form class="FormLabel" name="Login" id="Login" method="post" action="../../Login/Authenticate.asp">
         <input placeholder="Email Address" name="EmailAddress" width="10" type="text" /> <input width="10" placeholder="Password" name="Password" type="password" />
         <input type="submit" value="Log In" class="Button"> 
        </form>
         <a href="Information-Credentials.asp">Forgotten your username or password?</a>
          <br />
         <% Else 
		   SourceVar = "SELECT SUM(Items.UnitPrice * Basket.Quantity) AS Total, SUM(Basket.Quantity) AS Quantity,Basket.UserID FROM Basket INNER JOIN Items ON Basket.ItemID = Items.ItemID WHERE UserID = " & Session("UserID") & " GROUP BY Basket.UserID"
		   	Call OpenRecordSet(BasketInfo,0,SourceVar)
					While NOT BasketInfo.EOF
						TotalOrderValueVar = BasketInfo.Fields("Total")
						QuantityVar = BasketInfo.Fields("Quantity")
			BasketInfo.MoveNext()
			Wend
		BasketInfo.Close() %>
         
       <form name="LogOut" id="LogOut" action="../LogOut/LogOut-Process.asp">Logged in as <%= Session("FirstName") & " " & Session("Surname") %>  (<a href="../Profile/index.asp">Edit Profile</a>)<br />
        <a href="../Shop/Checkout/Basket.asp">Your Basket</a>: £<% If TotalOrderValueVar = "" Then %>0.00<% Else %><%= TotalOrderValueVar %>.00
       <% End If %> (<% If QuantityVar = 0 Then %>0 Items<% Else %>
	   <%= QuantityVar %> Item<% If RecordNumberVar <> 1 then %>s<% Else 
	   End If 
	   End If %>)<br />
       <input type="submit" value="Log Out" class="Button">
     </form></p>
    <% End If %>
   </div>
  </td>
 </tr>
</table>
 <table style="border: 1px thin; border-bottom:none; border-top: none; border-right:none; border-left:none;" cellpadding="0" border="0" width="100%" height="400" align="center">
  <tr align="center" valign="top">
   <td align="center" valign="top">
     <div id="tabs">
      <table cellpadding="10" width="1170" align="center" height="100%">
       <tr height="50" width="100%" valign="middle">
        <td width="1144" colspan="2" align="right" valign="top" style="border: 1px thin; border-top:none; border-bottom: 1px solid #000000; border-right:none; border-left:none;">  
          <div style="float:left;">
           <h1><a href="../Index.asp"><img src="../Images/Logo.png" width="300" height="103"/></a></h1>
          </div> 
         <p><br /><br /><br /><br /></p>
         <div id="title">
           <ul id="Menu">
           <li><a href="../../index.asp">home</a></li>
           <li><a href="../../Shop/Men.asp">men</a></li>
           <li><a href="../../Shop/Woman.asp">women</a></li>
           <li><a href="../../Shop/Accessories.asp">accessories</a></li>
           <li><a href="../../Shop/Popular.asp">popular</a></li>
           <% If Session("DenudeUser") = True then %>
           <li><a href="../../Management/Management.asp">management</a></li>
           <% Else 
              End If %>
          </ul>
        </div>
        </td>
       </tr>
      <tr>
      <td>
       <!-- InstanceBeginEditable name="MainContent" -->
        <div id="tabContent">
        <table border="0" width="100%" height="400">
         <tr>
          <td align="center" width="49%">&nbsp;</td>
          <td width="51%">
           <div align="right" style="float:right;">
           <form name="SignUp" method="post" action="../htdocs/Login-Authenticate.asp">
             <p id="FormLabel">
             <font size="+3" style="font-weight:lighter;">Sign Up</font> <br /><br />
             First Name  <input name="FirstName" type="text" /><br />
             Surname  <input name="Surname" type="text" /><br />
             Username  <input name="Username" type="text" /><br />
             Password  <input name="Password" type="password" /><br />
             Email Address  <input name="EmailAddress" type="text" /><br />
             </p>
              <input type="button" value="Sign Up" />
            </form>
            </div>
          </td>
         </tr>
        </table>
       </div>
      <!-- InstanceEndEditable -->
      </td>
     </tr>
    </table>
   </div>
 </td>
</tr>
<tr>
<td><p class="title">&nbsp;</p></td>
</tr>
<tr bgcolor="#FFFFFF" height="30" width="100%">
   <td width="1144" height="218" colspan="2" align="right" valign="top" style="border: 1px thin; border-bottom:none; border-top: 1px solid #000000; border-right:none; border-left:none;">
    <p align="center" id="FormLabel">&nbsp;</p><p align="center">Dènudè Ltd | 2012 | Copyright ©</p><p align="right" id="FormLabel"> FAQ's | Legal | Language | Contact | Help</p>
   </td>
  </tr>
 å</table>
</body>
<!-- InstanceEnd --></html>