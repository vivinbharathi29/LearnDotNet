<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
	<TITLE>Update Agency Status</TITLE>
<HEAD>
</HEAD>

	<FRAMESET ROWS="*,60" ID=TopWindow >
		<FRAME noresize frameborder=0 ID="UpperWindow" Name="UpperWindow" SRC="AgencyMain.asp?<%=Request.QueryString%>">
		<FRAME noresize frameborder=0 ID="LowerWindow" Name="LowerWindow" SRC="AgencyButtons.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>&<%=Request.QueryString%>">   
	</FRAMESET>
</HTML>