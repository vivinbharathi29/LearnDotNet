<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<!-- #include file = "../../includes/noaccess.inc" -->

<HTML>
<TITLE>View Lead Product Usage Exceptions</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*,60" ID=TopWindow >
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="ViewExceptionsMain.asp?PMID=<%=request("PMID")%>">
	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="ViewExceptionsButtons.asp">
</FRAMESET>

</HTML>