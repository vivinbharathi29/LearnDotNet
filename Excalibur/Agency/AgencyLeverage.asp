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
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="AgencyLeverageMain.asp?<%=Request.QueryString%>">
		<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="AgencyButtons.asp?<%=Request.QueryString%>">
	</FRAMESET>
</HTML>