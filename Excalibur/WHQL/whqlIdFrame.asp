<%@ Language=VBScript %>

<!-- #include file = "../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<HEAD>
<TITLE>WHQL ID Attachment</TITLE>
</HEAD>
	<FRAMESET ROWS="*" ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="whqlIdSubmit.aspx?<%=Request.QueryString%>">
	</FRAMESET>
</HTML>