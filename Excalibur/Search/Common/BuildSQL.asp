<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>



<HTML>
<HEAD>
	<TITLE>Build Advanced SQL Filters</TITLE>
</HEAD>

	<FRAMESET ROWS="*" ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="BuildSQLMain.asp?txtAdvanced=<%=Request("txtAdvanced")%>">
	</FRAMESET>
</HTML>