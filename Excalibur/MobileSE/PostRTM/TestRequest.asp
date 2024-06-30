<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<!-- #include file = "../../includes/noaccess.inc" -->
	
<%request("RootID")%>
<%request("ID")%>

<HTML>
	<TITLE>Sustaining Product Test</TITLE>
<HEAD>
</HEAD>

	<FRAMESET ROWS="160,30" ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="TestRequestMain.asp?RootID=<%=Request("RootID")%>&VersionID=<%=Request("ID")%>">
		<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="TestRequestButtons.asp">
	</FRAMESET>
</HTML>