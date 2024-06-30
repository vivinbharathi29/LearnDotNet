	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<TITLE>Preview</TITLE>
<FRAMESET ROWS="*" ID=TopWindow>
	<FRAME ID="ReportWindow" Name="ReportWindow" SRC="actionreportmain.asp?Action=<%=Request("Action")%>&ID=<%=Request("ID")%>&Type=<%=Request("Type")%>">
</FRAMESET>
