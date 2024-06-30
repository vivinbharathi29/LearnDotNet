<%@ Language=VBScript %>
<!-- #include file = "../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->
<HTML>
<HEAD>
<TITLE>Set Filter</TITLE>
</HEAD>
	<FRAMESET ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="FilterByCategory.aspx?<%=Request.QueryString%>">
	</FRAMESET>
</HTML>