<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>


<HTML>
<HEAD>
	<TITLE>Select Deliverables</TITLE>
</HEAD>
	<FRAMESET ROWS="*,60" ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="test.asp?ID=<%=request("ID")%>&ProdID=<%=request("ProdID")%>">
		<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="MapDeliverablesButtons.asp">
	</FRAMESET>
</HTML>