<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<TITLE>Add New Version</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*,60" ID=TopWindow >
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="SelectReleaseTypeMain.asp?ID=<%=request("ID")%>&RootID=<%=request("RootID")%>">
	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="SelectReleaseTypeButtons.asp">
</FRAMESET>
</HTML>