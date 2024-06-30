	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<HTML>
<HEAD>
<title>Switch PC</title>


</HEAD>
<FRAMESET ROWS="*" ID=TopWindow>
	<FRAME ID="MyWindow" Name="MyWindow" SRC="ChangePCMain.asp?<%=Request.QueryString %>">
</FRAMESET>
</HTML>
