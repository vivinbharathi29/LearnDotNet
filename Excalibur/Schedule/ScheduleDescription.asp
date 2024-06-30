<%@ Language=VBScript %>
<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
  
%>
<!-- #include file = "../includes/noaccess.inc" -->

<html>
<title>Edit Schedule Name</title>
</head>
<frameset ROWS="*" ID="TopWindow">
	<frame noresize ID="MainWindow" Name="MainWindow" SRC="../Schedule/SetScheduleDescription.asp?<%=Request.QueryString%>">
</frameset>
</html>