<%@  language="VBScript" %>
<%
	Response.Buffer = True
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
%>
<html>
<head>
	<title>Page Layout</title>
</head>
<frameset rows="*" id="TopWindow">
	<frame noresize="noresize" id="UpperWindow" src="PageLayoutMain.asp">
</frameset>
</html>
