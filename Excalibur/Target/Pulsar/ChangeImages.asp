<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<TITLE>Select Images</TITLE>
<HEAD>

</HEAD>
<%if Request("Type") = "1" then%>
<FRAMESET ROWS="*" ID=TopWindow >
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="ChangeImagesMain.asp?ProductID=<%=Request("ProductID")%>&RootID=<%=Request("RootID")%>&Type=<%=Request("Type")%>&VersionID=<%=Request("VersionID")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
</FRAMESET>
<%else%>
<FRAMESET ROWS="*,60" ID=TopWindow >
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="ChangeImagesMain.asp?ProductID=<%=Request("ProductID")%>&RootID=<%=Request("RootID")%>&Type=<%=Request("Type")%>&VersionID=<%=Request("VersionID")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="ChangeImagesButtons.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
</FRAMESET>
<%end if%>
</HTML>