<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<HEAD>
<TITLE>Support Category</TITLE>

</HEAD>
    <FRAMESET ROWS="*,65" ID=TopWindow >
	    <FRAME noresize ID="MainWindow" Name="MainWindow" SRC="CategoryMain.asp?ID=<%=request("ID")%>">
	    <FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="CategoryButtons.asp">
    </FRAMESET>
</HTML>