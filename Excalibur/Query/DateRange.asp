<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<TITLE>Date Range</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*" ID=TopWindow >
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="DateRangeMain.asp?StartDate=<%=request("StartDate")%>&EndDate=<%=request("EndDate")%>">
</FRAMESET>

</HTML>