<%@  language="VBScript" %>
<%
	Response.Buffer = True
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
%>
<html>
<head>
	<title>Custom Reports</title>
</head>
<frameset rows="*" id="TopWindow">
	<% if LCASE(Session("LoggedInUser")) = "auth\mahamilton" then %>
	<frame noresize="noresize" id="UpperWindow" src="CustomReportMain_mattH.asp?ProfileID=<%=request("ProfileID")%>&RunReportOK=<%=request("RunReportOK")%>">
	<% else %>
	<frame noresize="noresize" id="UpperWindow" src="CustomReportMain.asp?ProfileID=<%=request("ProfileID")%>&RunReportOK=<%=request("RunReportOK")%>">
	<% end if %>
</frameset>
</html>
