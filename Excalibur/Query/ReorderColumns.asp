<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>



<HTML>
	<TITLE>Reorder Columns</TITLE>
<HEAD>
</HEAD>

	<FRAMESET ROWS="*" ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="ReorderColumnsMain.asp?lstColumns=<%=Request("lstColumns")%>&UserSettingsID=<%=Request("UserSettingsID")%>">
	</FRAMESET>
</HTML>