<%@ Language=VBScript %>

<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
%>


<html>
<title>Deliverable Lookup</title>
<head>

</head>
<frameset rows="*,55" id="TopWindow">
	<frame id="UpperWindow" name="UpperWindow" src="FindSpareKitMain.asp?category=<%=Request.QueryString("Category") %>&PVID=<%=Request.QueryString("PVID") %>">
	<frame noresize id="LowerWindow" name="LowerWindow" src="FindSpareKitButtons.asp" scrolling="no">
</frameset>
<body />
</html>