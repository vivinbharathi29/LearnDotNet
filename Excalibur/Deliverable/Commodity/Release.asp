<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<!-- #include file = "../../includes/noaccess.inc" -->

<HTML>
	<TITLE>Release Versions</TITLE>
<HEAD>
</HEAD>

	<FRAMESET ROWS="*,60" ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="ReleaseMain.asp?VersionID=<%=Request("VersionID")%>&txtFunction=<%=request("txtFunction")%>">
		<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="ReleaseButtons.asp">
	</FRAMESET>
</HTML>
