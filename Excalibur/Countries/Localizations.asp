<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
	<TITLE>Select Localizations</TITLE>
	<FRAMESET ROWS="*,60" ID=TopWindow >
		<FRAME noresize frameborder="0" ID="UpperWindow" Name="UpperWindow" SRC="LocalizationsMain.asp?ID=<%=Request("ID")%>&FusionRequirements=<%=Request("FusionRequirements")%>&pvID=<%=Request("pvID")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
		<FRAME noresize frameborder="0" ID="LowerWindow" Name="LowerWindow" SRC="LocalizationsButtons.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
	</FRAMESET>
</HTML>