<%@ Language=VBScript %>
<!-- #include file = "../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->
<HTML>
<HEAD>
<TITLE>Choose a Market(s) and Plants(s) to Work With:</TITLE>
</HEAD>
	<FRAMESET ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="FilterByMktCampaignsAVSelector.aspx? <%=Request.QueryString%>">
	</FRAMESET>
</HTML>