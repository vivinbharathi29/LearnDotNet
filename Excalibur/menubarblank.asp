<!-- MenuBar.asp  -->
<%
strServername = LCase(Request.ServerVariables("SERVER_NAME"))
strURL = LCase(Request.ServerVariables("URL"))	
if left(strurl ,1) = "/" then
	strurl = mid(strurl,2)
end if
Select Case strServername
	Case "dwhorton3"
		strServer2 = "dwhorton3"
		strServername = strServername & "/ProgramOffice"
		strURL = mid(strURL, 16)
		strFileServer = "16.81.19.70"
	Case "localhost"
		strServer2 = "localhost"
		strServername = strServername & "/ProgramOffice"
		strURL = mid(strURL, 16)
		strFileServer = "16.81.19.70"
	Case Else
		strServer2 = strServername
		strFileServer = "dwhorton4"
end select

%>
<p><table cellPadding="1" border="0" cellSpacing="1" width="100%">
  
  <tr>
    <td nowrap width="180" style="VERTICAL-ALIGN: top" id="side">
</td>

    <td style="VERTICAL-ALIGN: top">
