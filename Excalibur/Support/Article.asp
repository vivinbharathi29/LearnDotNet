<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<TITLE>Support Article</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*,65" ID=TopWindow >
	<FRAME noresize ID="MainWindow" Name="MainWindow" SRC="ArticleMain.asp?ID=<%=request("ID")%>">
	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="ArticleButtons.asp">
</FRAMESET>

</HTML>