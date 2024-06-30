<!-- #include file = "../includes/noaccess.inc" -->
<!-- MenuBar.asp  -->
<%
Dim strExcaliburLink
strServername = LCase(Request.ServerVariables("SERVER_NAME")) & "/mobilese"
strURL = LCase(Request.ServerVariables("URL"))	
if left(strurl ,1) = "/" then
	strurl = mid(strurl,2)
end if
Select Case strServername
	Case Application("Excalibur_ServerName")
		strServer2 = Application("Excalibur_ServerName")
		strServername = strServername 
		strFileServer = Application("Excalibur_File_Server") & "\SE_WEB"
	Case "localhost"
		strServer2 = "localhost"
		strServername = strServername
		strURL = mid(strURL, 16)
		strFileServer = Application("Excalibur_File_Server") & "\SE_WEB"
	Case Else
		strServer2 = strServername
		strFileServer = Application("Excalibur_File_Server") & "\SE_WEB"
end select

'--Create link using current server name: ---
strExcaliburLink = "http://" & strServer2 & "/excalibur/Excalibur.asp"
%>
<p><table cellPadding="1" border="0" cellSpacing="1" width="100%">
  
  <tr>
    <td nowrap width="180" style="VERTICAL-ALIGN: top" id="side">
        <div style="display:none">
      <h1>
      Program Office<br>
      
<% If trim(lcase(strURL)) = "mobilese/default.asp" then %>     
      <img height="12" src="images/greyball.gif" width="12" align="middle"><b>Home</b><br>
<% Else %>  
     <img height="12" src="images/greyball.gif" width="12" align="middle"><a href="http://<%= strServerName%>/">Home</a><br>
<% End If %>
    <img height="12" src="images/greyball.gif" width="12" align="middle"><a target=newwindow href="file://<%= strFileServer%>/">Documents</a><br>
<BR>
</div>
	  Links<br>
      <img height="12" src="images/greyball.gif" width="12" align="middle"><a href="javascript:void(0);"  onclick="parent.location='../Excalibur.asp'">Pulsar</a><br> <!--Open Excalibur in parent window-->
      <img height="12" src="images/greyball.gif" width="12" align="middle"><a href="http://houcmitrel02.auth.hpicorp.net:81/mobilerelease.aspx">Release Lab</a><br>
      <img height="12" src="images/greyball.gif" width="12" align="middle"><a href="file://\\<%=Application("Excalibur_File_Server")%>\se_web\RestoreCDs\RestoreCD_Guide.xls">Restore CD Info</a><br>
      <br> 
<%	CurrentUser = lcase(Session("LoggedInUser"))%>

	  Reports<br>
	  <% If trim(lcase(strURL)) = "mobilese/products.asp" then %>     
			<img height="12" src="images/greyball.gif" width="12" align="middle"><b>Product Reports</b><br>
	  <%else%>

			<img height="12" src="images/greyball.gif" width="12" align="middle"><a href="products.asp">Product Reports</a><br>
      <%end if%>
    <BR>
     
  <BR>     Suggestions<br>
      <img height="12" src="images/greyball.gif" width="12" align="middle"><A HREF="mailto:max.yu@hp.com?Subject=PSE Website Process">Process</a><br>
      <img height="12" src="images/greyball.gif" width="12" align="middle"><A HREF="mailto:max.yu@hp.com?Subject=PSE Website">Web Site</A><br>
      </h1>
</td>
      

    <td style="VERTICAL-ALIGN: top">
