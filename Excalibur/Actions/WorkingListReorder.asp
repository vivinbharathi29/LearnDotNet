<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<!-- #include file = "../includes/noaccess.inc" -->


<HTML>
	<TITLE>Reorder Tasks</TITLE>
<HEAD>
</HEAD>

	<FRAMESET ROWS="*,60" ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="WorkingListReorderMain.asp?ID=<%=Request("ID")%>&ProjectID=<%=Request("ProjectID")%>&ReportOption=<%=Request("ReportOption")%>">
		<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="WorkingListReorderButtons.asp">
	</FRAMESET>
</HTML>