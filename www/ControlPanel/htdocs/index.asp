<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Session("BackgroungColour") = "#CCC"
   Session("CompanyInfoID") = 1
   Session("AuthenticationID") = 1
   Session("CompanyName") = "IndiCater Ltd"
   Session("FirstName") = "Ryan"
   Session("Surname") = "Fletcher"
   Session("JobTitle") = "Database Manager"
   Session("ControlPanelAdmin") = True
%>
<!--#include file="../Connections/Connections.asp" -->
<!--#include file="../Functions/PrivateFunctions.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head></head><!-- InstanceBegin template="/Templates/ControlPanelTemplate.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<!-- InstanceBeginEditable name="doctitle" -->
<title><%= Session("SystemName") %></title>
<!-- InstanceEndEditable -->
<style>
div#ContentHolder { border: 1px #CCC solid; box-shadow: 1px #999999; }
ul# { text-decoration:none; }
</style>
</head>
<!-- InstanceBeginEditable name="body" -->
<body bgcolor="<%= Session("BackgroundColour") %>">
<table align="center" cellpadding="10" cellspacing="10" width="100%">
<tr>
   <td colspan="2" height="102" align="center" valign="middle"><img width="300" height="100" src="file:///Macintosh%20HD/Users/RyanFletcher/Documents/Dénudé/Customisation/<%= Session("CompanyInfoID") %>/LogoV4.gif" /></td>
  </tr>
</table>  
<table align="center" cellpadding="10" cellspacing="10" width="100%" height="100%">
  <tr>
   <td valign="top" height="915" width="75%">
	<table height="40%" width="100%" border="1" align="center">
     <tr>
      <td valign="top" height="40%" width="40%">
       <div id="ContentHolder">
        <table>
         <tr>
          <td valign="top" width="177" rowspan="3"><img width="177" height="130" src="Customisation/<%= Session("CompanyInfoID") %>/<%= Session("AuthenticationID") %>.gif" /></td>
          <td valign="top" width="349">Name: <%= Session("FirstName") & " " & Session("Surname") %></td>
         </tr>
         <tr>
          <td valign="top">Company: <%= Session("CompanyName") %></td>
         </tr>
         <tr>
          <td valign="top" height="140">Job Title: <%= Session("JobTitle") %></td>
         </tr>
         <tr>
          <td valign="top"></td>
         </tr>  
        </table>
       </div>
      </td>
      <td valign="top" height="40%"width="60%">
       <div id="ContentHolder">
        <table>
         <% SourceVar  = "SELECT * FROM Clients"
		 	Call OpenRecordSet(Clients,1,SourceVar)
		 	While NOT Clients.EOF
			ClientNameVar = Clients.Fields("ClientName")
			%>
         <tr>
         <td width="519" height="48" valign="top"><p><% ClientNameVar %></p></td>
         </tr>
         <% Clients.MoveNext()
		 Wend
		Clients.Close() %>	 
         <tr>
          <td height="55" valign="top">System Notices:<% 'Call GetSystemNotices(CompanyInfoID,AuthenticationID) %></td>
         </tr>
         <tr>
          <td height="80" valign="top">Your Tasks:<% 'Call GetUserTasks(CompanyInfoID,AuthenticationID) %></td>
         </tr>
        </table>
      
       </div>
      </td>
     </tr>
    </table>
	<table height="30%" width="100%" border="1" align="center">
     <tr>
      <td width="100%">
       <div id="ContentHolder">
      
       </div>
      </td>
     </tr>
    </table>
    <table height="30%" width="100%" border="1" align="center">
     <tr>
      <td width="33.3%">
       <div id="ContentHolder">
      
       </div>
      </td>
      <td width="33.3%">
       <div id="ContentHolder">
      
       </div>
      </td>
      <td width="33.3%">
       <div id="ContentHolder">
      
       </div>
      </td>
     </tr>
    </table>
   </td>
   <td rowspan="2" valign="top" height="100%" width="25%">
    <table height="100%" width="100%" border="1" align="center">
     <tr>
      <td height="1260">
      <div align="center" id="ContentHolder">
       <table width="100%" height="100%" align="center" class="Modules">
<% 		 If Session("CompanyIntranet") = True Then %>
        <tr align="center">
      	 <td>Company</td>
        </tr>
<%       End If
		 If Session("Content") = True Then %> 
        <tr align="center">
         <td>Content</td>
        </tr>
<%       End If
         If Session("HR") = True Then %> 
        <tr align="center">
         <td>HR</td>
        </tr>
<%       End If
         If Session("Outlets") = True Then %>
        <tr align="center">
         <td>Outlets</td>
        </tr>
<%       End If
         If Session("UserAccess") = True Then %>
        <tr align="center">
         <td>User Access</td>
        </tr>
<%       End If
         If Session("Finance") = True Then %>
        <tr align="center">
         <td>Finance</td>
        </tr>
<%       End If
         If Session("Purchasing") = True Then %>
        <tr align="center">
         <td>Purchasing</td>
        </tr>
<%       End If 
         If Session("CRM") = True Then %>
        <tr align="center">
         <td>CRM</td>
        </tr>
<%       End If
         If Session("RecipeManager") = True Then %>
        <tr align="center">
         <td>Recipe Manager</td>
        </tr>
<%       End If
         If Session("Surveys") = True Then %>
        <tr align="center">
         <td>Surveys</td>
        </tr>
<%       End If %>
       </table>
      </ul> 
     </div>
    </td>
   </tr>
  </table>
 </td>
</tr>
<tr>
 <td height="46">
 <table align="center">
   <tr>
    <td align="center" valign="top"><p><strong><% Session("SystemName") %></strong><% Date() %></p></td>
   </tr>
  </table>
  </td>
  </tr>
 </table> 
</body>
<!-- InstanceEndEditable -->
<!-- InstanceEnd --></html>
