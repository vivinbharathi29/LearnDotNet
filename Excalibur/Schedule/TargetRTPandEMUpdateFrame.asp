<%@ Language=VBScript %>
<% Option Explicit %>
<%
	
Response.Buffer = True
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
	  
%>
<!-- #include file = "../includes/noaccess.inc" -->

<html>
<head>
<%	If Request("ID") <> "" Then %>
<title>Update Target RTP/MR and EM Dates</title>
</head>
        <frameset ROWS="*,60" ID="TopWindow">
	        <frame noresize ID="UpperWindow" Name="UpperWindow" SRC="TargetRTPandEMUpdateMain.asp?ID=<%=Request("ID")%>">
	        <frame noresize ID="LowerWindow" Name="LowerWindow" SRC="TargetRTPandEMUpdateButtons.asp">
        </frameset>
<%	Else %>
    <body>
        <font size="2" face="verdana"><br>Unable to display the requested page because not enough information was supplied.</font>
    </body>
<%	End If %>
</html>