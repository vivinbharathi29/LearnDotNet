<%@ Language=VBScript %>
<!-- #include file = "../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->
<HTML>
<HEAD>
<TITLE>Doc Kit Data</TITLE>
</HEAD>
	<FRAMESET ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="DocKitData.aspx?<%=Request.QueryString %>">
	</FRAMESET>
</HTML>