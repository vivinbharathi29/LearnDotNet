<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #include file = "../includes/noaccess.inc" -->

<%
	
Response.Buffer = True
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
	  
%>

<html>
<head>
<title>Schedule Batch Update</title>
</head>
<frameset ROWS="*" ID="TopWindow">
	<frame noresize ID="MainWindow" Name="MainWindow" SRC="../Schedule/MilestoneBatchUpdate_Pulsar.asp?<%=Request.QueryString%>">
</frameset>
</html>