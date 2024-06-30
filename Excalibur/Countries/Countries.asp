<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
	<TITLE>Select Countries</TITLE>
	<FRAMESET ROWS="*,60" ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="CountriesMain.asp?<%=Request.QueryString%>">
		<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="CountriesButtons.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
	</FRAMESET>
</HTML>