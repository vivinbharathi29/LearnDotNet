<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<TITLE>Support Articles</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*,57" ID=TopWindow >
	<FRAME noresize ID="MainWindow" Name="MainWindow" SRC="Preview.asp?List=1">
	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="ArticleListButtons.asp">
</FRAMESET>

</HTML>