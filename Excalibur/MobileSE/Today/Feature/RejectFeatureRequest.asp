<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<TITLE>Reject Feature Request</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*" ID=TopWindow >
	<FRAME noresize ID="MainWindow" Name="MainWindow" SRC="RejectFeatureRequestMain.asp?ID=<%=request("ID")%>">
</FRAMESET>

</HTML>