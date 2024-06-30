<%@  language="VBScript" %>
<%
	Response.Buffer = True
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
%>
<html>
<head>
	<title>Configure Report Section</title>
</head>
<frameset rows="*" id="TopWindow">
		<frame noresize="noresize" id="UpperWindow" src="ConfigureReportSectionsMain.asp?TypeID=<%=request("TypeID")%>&txtID=<%=request("txtID")%>&txtParams=<%=request("txtParams")%>">
	</frameset>
</html>
