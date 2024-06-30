<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<TITLE>Remove Root Deliverable</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*" ID=TopWindow >
	<FRAME noresize ID="MainWindow" Name="MainWindow" SRC="RemoveRootMain.asp?ProductID=<%=Request("ProductID")%>&DeliverableID=<%=Request("DeliverableID")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
</FRAMESET>

</HTML>