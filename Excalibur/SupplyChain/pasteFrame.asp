<%@ Language=VBScript %>

<!-- #include file = "../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<HEAD>
<TITLE>AV Import</TITLE>
</HEAD>
	<FRAMESET ROWS="*,60" ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="pasteDetail.asp?<%=Request.QueryString%>">
		<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="pasteButtons.asp">
	</FRAMESET>
</HTML>