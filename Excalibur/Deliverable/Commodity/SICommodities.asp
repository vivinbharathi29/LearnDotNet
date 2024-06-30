<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>


<HTML>
	<TITLE>Integration Test Commodities Received</TITLE>
<HEAD>
</HEAD>

	<FRAMESET ROWS="*,60" ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="SICommoditiesMain.asp?ID=<%=Request("ID")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
		<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="SICommoditiesButtons.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
	</FRAMESET>
</HTML>