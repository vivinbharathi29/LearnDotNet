<%@ Language=VBScript %>
	
	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
	
<HTML>
	<TITLE>Application Error</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*,55" ID=TopWindow>
	<FRAME ID="UpperWindow" Name="UpperWindow" SRC="AppErrorMain.asp?ID=<%=Request("ID")%>&CopyTo=<%=Request("CopyTo")%>">
	<FRAME ID="LowerWindow" Name="LowerWindow" SRC="AppErrorButtons.asp?ID=<%=Request("ID")%>" scrolling=no>
</FRAMESET>
</HTML>