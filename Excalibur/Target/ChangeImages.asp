<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<html>
<title>Select Images</title>
<head>
</head>
<%if Request("Type") = "1" then%>
<frameset rows="*" id="TopWindow">
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="ChangeImagesMain.asp?ProductID=<%=Request("ProductID")%>&RootID=<%=Request("RootID")%>&Type=<%=Request("Type")%>&VersionID=<%=Request("VersionID")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
</frameset>
<%else%>
<frameset rows="*,60" id="TopWindow">
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="ChangeImagesMain.asp?ProductID=<%=Request("ProductID")%>&RootID=<%=Request("RootID")%>&Type=<%=Request("Type")%>&VersionID=<%=Request("VersionID")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="ChangeImagesButtons.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
</frameset>
<%end if%>
</html>