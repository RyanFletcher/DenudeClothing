<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../../Connections/Connections.asp" -->
<!--#include file="../../Functions/PrivateFunctions.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>Control Panel | Login</title>
<link rel="stylesheet" href="../CSS/ControlPanel.css" />
</head>
<body>
<table  width="100%" height="100%">
 <tr>
  <td valign="middle">
   <table align="center">
    <tr>
 	 <td valign="middle">
      <form name="Login" method="post" action="Login-Authenticate.asp">
       <table border="1" cellpadding="10" cellspacing="10">
        <tr>
         <td align="center"><h1>Control Panel</h1></td>
        </tr>
        <tr>
         <td id="Borders" align="center">
         <input type="text" id="Username" name="Username" placeholder="Username" />
         <br />
         <% If Request.QueryString("InvalidUsername") = 1 Then Response.Write("Invalid Username") %>
         </td>
        </tr>   
        <tr>
         <td id="Borders" align="center">
         <input type="password" id="Password" name="Password" placeholder="Password" />
         <br />
         <% If Request.QueryString("InvalidPassword") = 1 Then Response.Write("Invalid Password") %>
         </td>
       </tr>
       <tr>
        <td align="center"><button type="submit">Login</button>
       </tr> 
       <tr>
       	<td align="left"><a onclick="ShowHide();" href="#">Forgotten your credentials?</a></td>
       </tr>  
      </table>
     </form> 
    </td>
   </tr>
  </table> 
 </td>
</tr>
</table>              
</body>
</html>
