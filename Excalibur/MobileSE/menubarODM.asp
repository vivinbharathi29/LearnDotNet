<!-- MenuBar.asp  -->
<%

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


%>
<p><table cellPadding="1" border="0" cellSpacing="1" width="100%">
  
  <tr>
    <td nowrap width="180" style="VERTICAL-ALIGN: top" id="side">
    </H1>
      <h1>
      Web Sites<br>
      
<% If trim(lcase(strURL)) = "mobilese/default.asp" then %>     
      <img height="12" src="images/greyball.gif" width="12" align="middle"><b>Home</b><br>
<% Else %>  
      <img height="12" src="images/greyball.gif" width="12" align="middle"><a href="http://<%= strServerName%>/">Home</a><br>
<% End If %>

 
 <BR>    
      Suggestions<br>
      <img height="12" src="images/greyball.gif" width="12" align="middle"><a href="mailto:max.yu@hp.com?Subject=Processe Suggestion">Process</a><br>
      <img height="12" src="images/greyball.gif" width="12" align="middle"><A HREF="mailto:max.yu@hp.com?Subject=SE Web Site Suggestion">Web Site</A><br>
      </h1>
</td>
      

    <td style="VERTICAL-ALIGN: top">
