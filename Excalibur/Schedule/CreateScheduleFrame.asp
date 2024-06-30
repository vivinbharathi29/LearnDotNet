<%@ Language=VBScript %>
<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
  
%>
<!-- #include file = "../includes/noaccess.inc" -->

<html>
<title>Create Schedule</title>
</head>
<frameset ROWS="*" ID="TopWindow">
	<frame noresize ID="MainWindow" Name="MainWindow" SRC="../Schedule/CreateScheduleMain.asp?<%=Request.QueryString%>">
</frameset>
</html>