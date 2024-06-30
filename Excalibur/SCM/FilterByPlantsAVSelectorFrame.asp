<%@ Language=VBScript %>
<!-- #include file = "../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->
<HTML>
<HEAD>
<TITLE>Choose Plant(s) and Product(s) to Work With:</TITLE>
</HEAD>
	<FRAMESET ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="FilterByPlantsAVSelector.aspx? <%=Request.QueryString%>">
	</FRAMESET>
</HTML>