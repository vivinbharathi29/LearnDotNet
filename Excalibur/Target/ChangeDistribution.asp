<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<TITLE>Change Distribution</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*" ID=TopWindow >
	<FRAME noresize ID="MainWindow" Name="MainWindow" SRC="ChangeDistributionMain.asp?ProductID=<%=Request("ProductID")%>&VersionID=<%=Request("VersionID")%>&RootID=<%=Request("RootID")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
</FRAMESET>

</HTML>