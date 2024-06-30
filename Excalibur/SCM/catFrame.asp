<%@ Language=VBScript %>

<!-- #include file = "../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<HEAD>
<TITLE>SCM AV Category Details</TITLE>
</HEAD>
	<FRAMESET ROWS="*,60" ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="catDetail.asp?<%=Request.QueryString%>">
		<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="catButtons.asp">
	</FRAMESET>
</HTML>