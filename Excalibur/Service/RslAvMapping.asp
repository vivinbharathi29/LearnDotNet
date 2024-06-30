<%@  language="VBScript" %>
<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
%>
<html>
<head>
    <title>Spare Kit Map Details</title>
</head>
<frameset rows="*" id="TopWindow">
<frame id="UpperWindow" name="UpperWindow" src="RSLAVMapping.aspx?<%= Request.QueryString %>">
</frameset>
</html>
