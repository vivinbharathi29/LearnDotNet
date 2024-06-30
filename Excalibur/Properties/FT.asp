<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>


<HTML>
<TITLE>Deliverable Changes</TITLE>
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=8" />
</HEAD>
<FRAMESET ROWS="*,55" ID=TopWindow >
	<FRAME noresize ID="MainWindow" Name="MainWindow" SRC="FTMain.asp?ID=<%=Request("ID")%>&RootID=<%=Request("RootID")%>&app=<%=Request("app")%>">
	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="FTButtons.asp?app=<%=Request("app")%>"> 
</FRAMESET>

</HTML>