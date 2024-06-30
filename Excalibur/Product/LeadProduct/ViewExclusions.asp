<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<!-- #include file = "../../includes/noaccess.inc" -->

<HTML>
<TITLE>View Lead Product Exclusions</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*,60" ID=TopWindow >
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="ViewExclusionsMain.asp?PMID=<%=request("PMID")%>">
	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="ViewExclusionsButtons.asp">
</FRAMESET>

</HTML>