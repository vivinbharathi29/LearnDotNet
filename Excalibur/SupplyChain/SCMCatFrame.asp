<%@ Language=VBScript %>

<!-- #include file = "../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<HEAD>
<TITLE>SCM Category Details</TITLE>
</HEAD>
	<FRAMESET ROWS="*,60" ID=TopWindow >
		<FRAME noresize ID="SCMUpperWindow" Name="SCMUpperWindow" SRC="SCMCatDetail.asp?<%=Request.QueryString%>">
		<FRAME noresize ID="SCMLowerWindow" Name="SCMLowerWindow" SRC="SCMCatButtons.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
	</FRAMESET>
</HTML>