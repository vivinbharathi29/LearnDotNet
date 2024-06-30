<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<!-- #include file = "../includes/noaccess.inc" -->


<HTML>
	<TITLE>Reorder Roadmap Items</TITLE>
<HEAD>
</HEAD>

	<FRAMESET ROWS="*" ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="ScheduleReorderMain.asp?ID=<%=Request("ID")%>">
	</FRAMESET>
</HTML>